using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// Puts one page's worth of cells on the page: what is painted behind them, the lines around
/// them, and the text inside them.
/// </summary>
/// <remarks>
/// The three passes are separate because they layer: a fill would cover the border of the cell
/// above it, and a border would cover the text of the cell before it. Excel draws them in this
/// order and so does this.
/// </remarks>
internal sealed partial class CellPlacer
{
    private readonly PdfExportContext _context;
    private readonly Worksheet _sheet;
    private readonly GridMeasurer _grid;
    private readonly ComposedPage _page;
    private readonly PageSlots _columns;
    private readonly PageSlots _rows;
    private readonly Dictionary<(int Row, int Column), RangeRef> _merges = [];
    private readonly HashSet<(int Row, int Column)> _covered = [];
    private readonly Dictionary<(int Row, int Column), ConditionalAppearance> _conditionalByCell = [];

    internal CellPlacer(
        PdfExportContext context,
        Worksheet sheet,
        GridMeasurer grid,
        ComposedPage page,
        PageSlots columns,
        PageSlots rows)
    {
        _context = context;
        _sheet = sheet;
        _grid = grid;
        _page = page;
        _columns = columns;
        _rows = rows;

        foreach (RangeRef merge in sheet.MergedRanges)
        {
            if (!Reaches(merge))
                continue;

            _merges[(merge.Start.Row, merge.Start.Column)] = merge;
            Swallow(merge);
        }
    }

    /// <summary>
    /// The structure elements of the sheet, when the export is tagged: the sheet is a table, a
    /// row of it is a row, and a cell is a header cell when it is one of the repeated titles.
    /// </summary>
    public SheetTags? Tags { get; init; }

    /// <summary>The interactive comment anchors of the sheet, when the caller opted into them.</summary>
    public SheetCommentPlan? Comments { get; init; }

    /// <summary>
    /// How far right the page itself reaches, in the grid's own space. Text spilling out of the
    /// last column stops here rather than running off the paper.
    /// </summary>
    public double PageReach { get; init; } = double.MaxValue;

    /// <summary>Whether any part of a merge falls on this page.</summary>
    private bool Reaches(RangeRef merge) =>
        _rows.Meets(merge.Start.Row, merge.End.Row) && _columns.Meets(merge.Start.Column, merge.End.Column);

    /// <summary>
    /// Notes the cells a merge swallows — the ones it covers other than the corner it is anchored
    /// to, which is the only one that draws anything.
    /// </summary>
    /// <remarks>
    /// Only the cells on this page are noted. A merge is a rectangle of any size a sheet allows,
    /// and a whole column of one is a legal merge: writing out every cell of every merge would
    /// cost the area of the merges rather than the area of the paper, which for a large one is
    /// the difference between a page and a million of them.
    /// </remarks>
    private void Swallow(RangeRef merge)
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            int row = _rows.Lines[r];
            if (row < merge.Start.Row || row > merge.End.Row)
                continue;

            for (int c = 0; c < _columns.Count; c++)
            {
                int column = _columns.Lines[c];
                if (column < merge.Start.Column || column > merge.End.Column)
                    continue;

                if (row != merge.Start.Row || column != merge.Start.Column)
                    _covered.Add((row, column));
            }
        }
    }

    /// <summary>The style of one cell, which every pass needs and none should look up twice.</summary>
    /// <summary>How the sheet's conditional formatting paints its cells.</summary>
    public ConditionalFormatEvaluator? Conditional { get; init; }

    /// <summary>
    /// The style a cell is actually drawn in: the one it states, with whatever conditional
    /// formatting matched laid over it.
    /// </summary>
    /// <remarks>
    /// One place, because every pass has to agree. A rule that changes the number format changes
    /// how wide the text is, so the pass that measures and the pass that draws must both see it;
    /// a fill pass that knew about the rules and a border pass that did not would paint a cell
    /// half-conditioned.
    /// </remarks>
    private CellStyle StyleOf(int row, int column)
    {
        CellStyle stated = _sheet.Workbook.GetStyle(new StyleId(_sheet.StyleIndexes.Get(row, column)));
        return Conditional is { HasRules: true } evaluator
            ? evaluator.Overlay(stated, ConditionalOf(row, column))
            : stated;
    }

    /// <summary>What the rules come to for a cell, or nothing when the sheet has none.</summary>
    /// <remarks>
    /// Remembered for the page rather than worked out again: the fills, the borders and the text
    /// are three passes over the same cells, and judging every rule three times over would be
    /// three times the work for one answer. The memo is the size of a page, not of a sheet.
    /// </remarks>
    private ConditionalAppearance ConditionalOf(int row, int column)
    {
        if (Conditional is not { HasRules: true } evaluator)
            return default;

        if (_conditionalByCell.TryGetValue((row, column), out ConditionalAppearance cached))
            return cached;

        ConditionalAppearance appearance = evaluator.Of(new CellRef(row, column));
        _conditionalByCell[(row, column)] = appearance;
        return appearance;
    }

    /// <summary>The element a cell's text belongs to, or null when the export is not tagged.</summary>
    private TagRef? Tag(int row, int column) => Tags?.Cell(row, column);

    /// <summary>Paints what sits behind the cells.</summary>
    public void Fills()
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            for (int c = 0; c < _columns.Count; c++)
            {
                int row = _rows.Lines[r];
                int column = _columns.Lines[c];
                if (_covered.Contains((row, column)))
                    continue;

                ConditionalAppearance conditional = ConditionalOf(row, column);
                FillStyle fill = StyleOf(row, column).Fill;

                bool graded = !conditional.ScaleColor.IsAuto;
                if (fill.IsEmpty && !graded && conditional.Bar is null && conditional.Icon is null)
                    continue;

                (double x, double y, double width, double height) = Box(c, r, row, column);

                // A colour scale grades the cell itself, so it stands in for whatever fill the
                // cell states rather than sitting on top of it.
                ExcelColor paint = graded
                    ? conditional.ScaleColor
                    : fill.Gradient is { Stops.Count: > 0 } gradient ? gradient.Stops[0].Color : fill.Color;

                if (PdfExportContext.HasColor(paint))
                {
                    if (!graded)
                        WarnAboutFill(fill);

                    _page.Items.Add(new FillItem
                    {
                        X = x,
                        Y = y,
                        Width = width,
                        Height = height,
                        Color = _context.ColorOf(paint, PdfColor.White),
                    });
                }

                if (conditional.Bar is { } bar)
                    DrawBar(bar, x, y, width, height);

                if (conditional.Icon is { } icon)
                    DrawIcon(icon, x, y, height);
            }
        }
    }

    /// <summary>Says what a fill loses on the way to the page, once per kind.</summary>
    private void WarnAboutFill(FillStyle fill)
    {
        if (fill.Gradient is not null)
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.LayoutApproximated,
                "A gradient fill is painted flat in the colour of its first stop.",
                "gradient-fill");
        }
        else if (fill.Pattern != FillPattern.Solid && fill.Pattern != FillPattern.None)
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.LayoutApproximated,
                "A patterned fill is painted in the colour of its pattern rather than woven.",
                "pattern-fill");
        }
    }

    /// <summary>
    /// Draws the bar of a data-bar rule: a band from the left edge of the cell, as long a share
    /// of the cell as the value is of the range.
    /// </summary>
    private void DrawBar(ConditionalBar bar, double x, double y, double width, double height)
    {
        double inset = Math.Min(2, height / 4);
        double room = width - (2 * inset);
        double length = room * bar.Fraction;
        if (length <= 0 || height <= 2 * inset)
            return;

        _page.Items.Add(new FillItem
        {
            X = x + inset,
            Y = y + inset,
            Width = length,
            Height = height - (2 * inset),
            Color = _context.ColorOf(bar.Color, PdfColor.FromRgb(0x638EC6)),
        });
    }

    /// <summary>Draws the icon of an icon-set rule, at the left edge where Excel puts it.</summary>
    private void DrawIcon(ConditionalIcon icon, double x, double y, double height)
    {
        if (IconSetShapes.Of(icon.SetName, icon.Index, icon.Count) is not { } drawn)
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.ContentSkipped,
                $"The icon set '{icon.SetName}' is not one this draws; the cell is printed without an icon.",
                $"icon-set-{icon.SetName}");
            return;
        }

        if (drawn.Shape == IconShape.None)
            return;

        double size = Math.Min(IconSize(height), height);
        if (size <= 0)
            return;

        _page.Items.Add(new IconItem
        {
            X = x + Padding,
            Y = y + ((height - size) / 2),
            Shape = drawn.Shape,
            Color = drawn.Color,
            Size = size,
            Angle = drawn.Angle,
            Fraction = drawn.Fraction,
            Parts = icon.Count,
        });
    }

    /// <summary>How big an icon is drawn: as tall as the row allows, within reason.</summary>
    private static double IconSize(double height) => Math.Min(11, height - 2);

    /// <summary>
    /// How much room an icon takes from the left of a cell, which the text then starts after.
    /// </summary>
    private double IconRoom(int row, int column, double height)
    {
        ConditionalAppearance conditional = ConditionalOf(row, column);
        if (conditional.Icon is not { } icon ||
            IconSetShapes.Of(icon.SetName, icon.Index, icon.Count) is not { } drawn ||
            drawn.Shape == IconShape.None)
            return 0;

        return Math.Max(0, Math.Min(IconSize(height), height)) + Padding;
    }

    /// <summary>Whether a rule hides the number and shows only its icon or bar.</summary>
    private bool HidesValue(int row, int column)
    {
        ConditionalAppearance conditional = ConditionalOf(row, column);
        return conditional.Icon is { ShowValue: false } || conditional.Bar is { ShowValue: false };
    }

    /// <summary>
    /// Where a cell sits on the page and how big it is, following a merge across whatever part
    /// of it this page shows.
    /// </summary>
    private (double X, double Y, double Width, double Height) Box(int c, int r, int row, int column)
    {
        double x = _columns.Near[c];
        double y = _rows.Near[r];
        double width = _columns.Size[c];
        double height = _rows.Size[r];

        if (!_merges.TryGetValue((row, column), out RangeRef merge))
            return (x, y, width, height);

        for (int line = column + 1; line <= merge.End.Column; line++)
        {
            int slot = _columns.SlotOf(line);
            if (slot < 0)
                break;

            width += _columns.Size[slot];
        }

        for (int line = row + 1; line <= merge.End.Row; line++)
        {
            int slot = _rows.SlotOf(line);
            if (slot < 0)
                break;

            height += _rows.Size[slot];
        }

        return (x, y, width, height);
    }
}
