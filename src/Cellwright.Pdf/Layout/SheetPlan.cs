namespace Cellwright.Pdf.Layout;

/// <summary>
/// What of a sheet is printed: the rectangle of cells, with the hidden rows and columns already
/// dropped and the repeated titles named.
/// </summary>
/// <remarks>
/// A hidden row is not printed and takes no room, so it is left out here rather than checked for
/// at every step afterwards. What remains is two lists of line numbers — the rows and the
/// columns that print, in order — which is what everything downstream walks.
/// </remarks>
internal sealed class SheetPlan
{
    private SheetPlan(int[] rows, int[] columns, RangeRef area)
    {
        Rows = rows;
        Columns = columns;
        Area = area;
    }

    /// <summary>The rows that print, in order, 1-based.</summary>
    public int[] Rows { get; }

    /// <summary>The columns that print, in order, 1-based.</summary>
    public int[] Columns { get; }

    /// <summary>The rectangle the rows and columns came from.</summary>
    public RangeRef Area { get; }

    /// <summary>The rows repeated at the top of every page, or an empty span.</summary>
    public int[] TitleRows { get; private init; } = [];

    /// <summary>The columns repeated at the left of every page.</summary>
    public int[] TitleColumns { get; private init; } = [];

    /// <summary>
    /// What a sheet prints, or <see langword="null"/> when it prints nothing at all: an empty
    /// sheet has no page in Excel either.
    /// </summary>
    /// <param name="context">The render in progress.</param>
    /// <param name="sheet">The sheet to plan.</param>
    public static SheetPlan? For(PdfExportContext context, Worksheet sheet)
    {
        RangeRef? printed = sheet.PrintArea ?? Used(sheet);
        if (printed is not { } area)
            return null;

        // A print area beyond the used range is honoured: the author asked for that much paper.
        int[] rows = Visible(area.Start.Row, area.End.Row, row => sheet.Row(row).Hidden);
        int[] columns = Visible(area.Start.Column, area.End.Column, column => sheet.Column(column).Hidden);

        if (rows.Length == 0 || columns.Length == 0)
            return null;

        (int First, int Last)? titleRows = sheet.RepeatRows;
        (int First, int Last)? titleColumns = sheet.RepeatColumns;

        return new SheetPlan(rows, columns, area)
        {
            TitleRows = titleRows is { } tr
                ? Visible(tr.First, tr.Last, row => sheet.Row(row).Hidden)
                : [],
            TitleColumns = titleColumns is { } tc
                ? Visible(tc.First, tc.Last, column => sheet.Column(column).Hidden)
                : [],
        };
    }

    /// <summary>
    /// What a sheet prints when it names no print area: the cells it uses, the merges anchored
    /// among them and the cells its pictures hang from.
    /// </summary>
    /// <remarks>
    /// A merge across three columns whose last two hold nothing still occupies them, and a
    /// picture pinned to a cell past the used range still prints. Both would otherwise fall
    /// outside the grid, and a heading would be squeezed into one narrow column while a picture
    /// vanished.
    /// </remarks>
    private static RangeRef? Used(Worksheet sheet)
    {
        int firstRow = int.MaxValue;
        int firstColumn = int.MaxValue;
        int lastRow = 0;
        int lastColumn = 0;

        if (sheet.UsedRange is { } used)
            Cover(used, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);

        foreach (RangeRef merge in sheet.MergedRanges)
            Cover(merge, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);

        foreach (Picture picture in sheet.Pictures)
        {
            if (picture.IsHidden)
                continue;

            RangeRef hung = picture.Anchor.Kind switch
            {
                PictureAnchorKind.Absolute => Floating(sheet, picture.Anchor),
                PictureAnchorKind.TwoCell => new RangeRef(picture.Anchor.From, picture.Anchor.To),
                _ => new RangeRef(picture.Anchor.From, picture.Anchor.From),
            };

            Cover(hung, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
        }

        // A chart extends the printed grid the way a picture does: the cells its frame hangs
        // over have to be on some page for the frame to be anywhere.
        foreach (Chart chart in sheet.Charts)
        {
            RangeRef hung = chart.Anchor.Kind switch
            {
                PictureAnchorKind.Absolute => Floating(sheet, chart.Anchor),
                PictureAnchorKind.TwoCell => new RangeRef(chart.Anchor.From, chart.Anchor.To),
                _ => new RangeRef(chart.Anchor.From, chart.Anchor.From),
            };

            Cover(hung, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
        }

        return lastRow == 0
            ? null
            : new RangeRef(new CellRef(firstRow, firstColumn), new CellRef(lastRow, lastColumn));
    }

    /// <summary>The cells an absolutely placed picture reaches.</summary>
    /// <remarks>
    /// <para>
    /// A picture pinned to no cell at all is still printed, and whether it is printed at all
    /// depends on whether the sheet prints the part of the grid it floats over — a sheet holding
    /// nothing but a picture would otherwise have no page to put it on.
    /// </para>
    /// <para>
    /// The rectangle runs from A1 rather than from the picture: an absolute position is measured
    /// from the corner of the sheet, so the corner has to be the corner of the printed grid too,
    /// or everything drawn from it lands a few columns out.
    /// </para>
    /// <para>
    /// The sizes are measured with the nominal digit of the standard font rather than the one the
    /// render goes on to use. The answer only says which cells are planned, and it takes a whole
    /// column of error to change that.
    /// </para>
    /// </remarks>
    private static RangeRef Floating(Worksheet sheet, PictureAnchor anchor)
    {
        int lastColumn = Reaches(
            (anchor.Position.X + anchor.Extent.Width) / SheetUnits.EmuPerPoint,
            SheetLimits.MaxColumn,
            column => SheetUnits.ColumnPoints(
                sheet.Column(column).Width ?? SheetUnits.DefaultColumnWidth, SheetUnits.NominalDigitPixels));

        int lastRow = Reaches(
            (anchor.Position.Y + anchor.Extent.Height) / SheetUnits.EmuPerPoint,
            SheetLimits.MaxRow,
            row => sheet.Row(row).Height ?? SheetUnits.DefaultRowHeight);

        return new RangeRef(new CellRef(1, 1), new CellRef(lastRow, lastColumn));
    }

    /// <summary>The line a run from the sheet's own edge ends within.</summary>
    /// <param name="to">Where the run ends, in points.</param>
    /// <param name="limit">The last line the sheet has.</param>
    /// <param name="sizeOf">How big one line is.</param>
    private static int Reaches(double to, int limit, Func<int, double> sizeOf)
    {
        double at = 0;

        for (int line = 1; line <= limit; line++)
        {
            at += sizeOf(line);
            if (at >= to)
                return line;
        }

        return limit;
    }

    /// <summary>Grows a rectangle to take in another.</summary>
    private static void Cover(RangeRef range, ref int firstRow, ref int firstColumn, ref int lastRow, ref int lastColumn)
    {
        firstRow = Math.Min(firstRow, range.Start.Row);
        firstColumn = Math.Min(firstColumn, range.Start.Column);
        lastRow = Math.Max(lastRow, range.End.Row);
        lastColumn = Math.Max(lastColumn, range.End.Column);
    }

    private static int[] Visible(int first, int last, Func<int, bool> hidden)
    {
        if (last < first)
            return [];

        List<int> kept = new(Math.Min(last - first + 1, 1024));
        for (int line = first; line <= last; line++)
        {
            if (!hidden(line))
                kept.Add(line);
        }

        return [.. kept];
    }
}
