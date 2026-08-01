using Cellwright.Styles;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// How wide every printed column is and how tall every printed row, in points.
/// </summary>
/// <remarks>
/// The widths come from the workbook's standard font: a column width is a count of digits, so
/// what it comes to on paper depends on how wide that font draws a digit. The heights come from
/// the file when it states them — Excel has already measured those — and from the tallest font
/// in the row when it does not.
/// </remarks>
internal sealed class GridMeasurer
{
    private readonly Worksheet _sheet;
    private readonly SheetPlan _plan;
    private readonly double[] _widths;
    private readonly double[] _heights;
    private readonly Dictionary<int, int> _rowAt = [];
    private readonly Dictionary<int, int> _columnAt = [];

    private double[]? _across;
    private double[]? _down;

    internal GridMeasurer(PdfExportContext context, Worksheet sheet, SheetPlan plan)
    {
        _sheet = sheet;
        _plan = plan;

        Appearances = new CellAppearances(context);
        double digit = SheetUnits.DigitPixels(Appearances.Standard.DigitWidth);

        _widths = new double[plan.Columns.Length];
        for (int i = 0; i < plan.Columns.Length; i++)
        {
            double width = sheet.Column(plan.Columns[i]).Width ?? SheetUnits.DefaultColumnWidth;
            _widths[i] = SheetUnits.ColumnPoints(width, digit);
            _columnAt[plan.Columns[i]] = i;
        }

        _heights = new double[plan.Rows.Length];
        for (int i = 0; i < plan.Rows.Length; i++)
        {
            _heights[i] = sheet.Row(plan.Rows[i]).Height ?? SheetUnits.DefaultRowHeight;
            _rowAt[plan.Rows[i]] = i;
        }
    }

    /// <summary>The appearances the sheet's cells are drawn in.</summary>
    public CellAppearances Appearances { get; }

    /// <summary>How wide the printed column at a position is, in points.</summary>
    /// <param name="index">The position among the printed columns.</param>
    public double WidthAt(int index) => _widths[index];

    /// <summary>How tall the printed row at a position is, in points.</summary>
    /// <param name="index">The position among the printed rows.</param>
    public double HeightAt(int index) => _heights[index];

    /// <summary>Raises a row's height, which is what an auto-sized wrapped cell does to it.</summary>
    /// <param name="index">The position among the printed rows.</param>
    /// <param name="height">The height it needs, in points.</param>
    public void GrowRow(int index, double height)
    {
        if (height <= _heights[index])
            return;

        _heights[index] = height;
        _down = null;
    }

    /// <summary>How far across the printed grid a column starts, in points from its left edge.</summary>
    /// <param name="column">The column number, 1-based, printed or not.</param>
    /// <remarks>
    /// A column the sheet does not print takes up no room, so it stands where the next printed
    /// one does; a column past them all stands at the far edge. That is what lets something
    /// pinned outside this page — a picture whose other corner is on the next one — still be
    /// placed against the part of the grid that is here.
    /// </remarks>
    public double ColumnPosition(int column) =>
        Position(_plan.Columns, _across ??= Prefix(_widths), column);

    /// <summary>How far down the printed grid a row starts, in points from its top edge.</summary>
    /// <param name="row">The row number, 1-based, printed or not.</param>
    public double RowPosition(int row) => Position(_plan.Rows, _down ??= Prefix(_heights), row);

    /// <summary>The running total of the sizes before each printed line, and the whole after them.</summary>
    private static double[] Prefix(double[] sizes)
    {
        var prefix = new double[sizes.Length + 1];
        for (int i = 0; i < sizes.Length; i++)
            prefix[i + 1] = prefix[i] + sizes[i];

        return prefix;
    }

    private static double Position(int[] lines, double[] prefix, int line)
    {
        int index = Array.BinarySearch(lines, line);

        // Not printed: the search says where it would go, which is the printed line that has
        // taken its place — or the end of the grid, when there is none after it.
        return prefix[index >= 0 ? index : Math.Min(~index, lines.Length)];
    }

    /// <summary>Whether a row states its own height, which auto-sizing must not override.</summary>
    /// <param name="index">The position among the printed rows.</param>
    public bool RowHeightIsStated(int index) => _sheet.Row(_plan.Rows[index]).Height is not null;

    /// <summary>Where a printed row sits among them, or -1 when it is not printed.</summary>
    /// <param name="row">The row number, 1-based.</param>
    public int RowIndex(int row) => _rowAt.TryGetValue(row, out int index) ? index : -1;

    /// <summary>Where a printed column sits among them, or -1 when it is not printed.</summary>
    /// <param name="column">The column number, 1-based.</param>
    public int ColumnIndex(int column) => _columnAt.TryGetValue(column, out int index) ? index : -1;

    /// <summary>How wide a run of printed columns is together.</summary>
    /// <param name="from">The first position.</param>
    /// <param name="count">How many columns.</param>
    public double WidthOf(int from, int count)
    {
        double total = 0;
        for (int i = from; i < from + count && i < _widths.Length; i++)
            total += _widths[i];

        return total;
    }

    /// <summary>How tall a run of printed rows is together.</summary>
    /// <param name="from">The first position.</param>
    /// <param name="count">How many rows.</param>
    public double HeightOf(int from, int count)
    {
        double total = 0;
        for (int i = from; i < from + count && i < _heights.Length; i++)
            total += _heights[i];

        return total;
    }
}

/// <summary>
/// The appearances of a sheet's cells, built once per distinct font style.
/// </summary>
/// <remarks>
/// A workbook interns its styles already, so the same font style arrives here as the same
/// record: value equality makes the cache key, and a sheet of a million cells embeds its
/// handful of faces once each.
/// </remarks>
internal sealed class CellAppearances
{
    private readonly PdfExportContext _context;
    private readonly Dictionary<FontStyle, TextAppearance> _cache = [];

    internal CellAppearances(PdfExportContext context)
    {
        _context = context;

        // The standard font is what a column width counts digits of: the one the default style
        // names, which is the style every cell inherits from.
        Standard = Of(context.Source.GetStyle(default).Font);
    }

    /// <summary>The appearance of the workbook's standard font.</summary>
    public TextAppearance Standard { get; }

    /// <summary>The appearance a font style draws in.</summary>
    /// <param name="font">The font a cell's style names.</param>
    public TextAppearance Of(FontStyle font)
    {
        ArgumentNullException.ThrowIfNull(font);

        if (_cache.TryGetValue(font, out TextAppearance? cached))
            return cached;

        TextAppearance appearance = Build(font);
        _cache[font] = appearance;
        return appearance;
    }

    private TextAppearance Build(FontStyle font)
    {
        Inkwright.Fonts.PdfFont face = _context.Fonts.Resolve(font);
        double size = font.Size > 0 ? font.Size : 11;

        // A superscript or subscript is drawn smaller and off the baseline, as Excel draws it.
        double drawn = font.VerticalAlign == FontVerticalAlign.Baseline ? size : size * 0.65;
        double rise = font.VerticalAlign switch
        {
            FontVerticalAlign.Superscript => size * 0.33,
            FontVerticalAlign.Subscript => -size * 0.14,
            _ => 0,
        };

        return new TextAppearance
        {
            Font = face,
            Size = drawn,
            Color = _context.ColorOf(font.Color, Inkwright.PdfColor.Black),
            Ascent = face.Ascent * size / 1000,
            Descent = -face.Descent * size / 1000,
            Underline = font.Underline,
            Strikethrough = font.Strikethrough,
            Rise = rise,
        };
    }
}
