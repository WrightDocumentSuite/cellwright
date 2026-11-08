namespace Cellwright.Pdf.Layout;

/// <summary>
/// Cutting a grid into pages. A spreadsheet paginates in two directions at once: the columns
/// are cut into bands that fit the width of the paper and the rows into bands that fit its
/// height, and every pairing of the two is a page.
/// </summary>
internal sealed partial class SheetComposer
{
    /// <summary>The smallest a sheet is ever drawn at, below which nothing could be read.</summary>
    private const double MinimumScale = 0.1;

    /// <summary>One run of consecutive printed lines that share a page.</summary>
    private readonly record struct Band(int From, int Count);

    private List<ComposedPage> Paginate(SheetPlan plan, GridMeasurer grid)
    {
        PageBudget.Check(_firstIndex + Least(plan, grid), _context.Options);

        MeasureAutoHeights(plan, grid);

        double scale = Scale(plan, grid);
        (List<Band> columns, List<Band> rows) = Bands(plan, grid, scale);

        PageBudget.Check(_firstIndex + (columns.Count * rows.Count), _context.Options);

        List<ComposedPage> pages = [];
        foreach ((Band columnBand, Band rowBand) in Order(columns, rows))
        {
            ComposedPage page = NewPage(pages.Count);
            page.Scale = scale;
            DrawBand(page, plan, grid, columnBand, rowBand);
            pages.Add(page);
        }

        return pages;
    }

    /// <summary>
    /// The fewest pages a sheet could come to, worked out before anything is measured.
    /// </summary>
    /// <remarks>
    /// Two things can still change the count afterwards, and both can only raise it: measuring
    /// grows the rows that size themselves, and the print scale is never smaller than a tenth.
    /// So the count at the largest room a page could ever leave is a floor, and a sheet already
    /// past its budget here is past it whatever happens next — which is worth knowing before the
    /// measuring pass walks every cell of it.
    /// </remarks>
    private int Least(SheetPlan plan, GridMeasurer grid)
    {
        double scale = _setup.FitToPage ? MinimumScale : Math.Clamp(_setup.Scale, 10, 400) / 100.0;
        (List<Band> columns, List<Band> rows) = Bands(plan, grid, scale);
        return columns.Count * rows.Count;
    }

    /// <summary>
    /// The bands a grid falls into at a scale. The room a page leaves is measured in the grid's
    /// own space, so a scale of a half means twice as much of the grid fits on the paper.
    /// </summary>
    private (List<Band> Columns, List<Band> Rows) Bands(SheetPlan plan, GridMeasurer grid, double scale)
    {
        double headingWidth = _setup.PrintHeadings ? RowHeadingBandWidth(plan, grid) : 0;
        double headingHeight = _setup.PrintHeadings ? ColumnHeadingBandHeight(grid) : 0;
        double width = (_geometry.ContentWidth / scale) - headingWidth - Titles(plan.TitleColumns, grid, horizontal: true);
        double height = (_geometry.ContentHeight / scale) - headingHeight - Titles(plan.TitleRows, grid, horizontal: false);

        return (
            Split(plan.Columns, index => grid.WidthAt(index), width, _setup.ColumnBreaks),
            Split(plan.Rows, index => grid.HeightAt(index), height, _setup.RowBreaks));
    }

    private static double RowHeadingBandWidth(SheetPlan plan, GridMeasurer grid)
    {
        int maxRow = plan.Rows.Length > 0 ? plan.Rows[^1] : 1;
        string widest = maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return Math.Max(22.0, grid.Appearances.Standard.Measure(widest) + 10.0);
    }

    private static double ColumnHeadingBandHeight(GridMeasurer grid) =>
        Math.Max(SheetUnits.DefaultRowHeight, grid.Appearances.Standard.LineHeight + 4.0);

    /// <summary>How much room the repeated titles take on every page of a band.</summary>
    private static double Titles(int[] titles, GridMeasurer grid, bool horizontal)
    {
        double total = 0;
        foreach (int line in titles)
        {
            int index = horizontal ? grid.ColumnIndex(line) : grid.RowIndex(line);
            if (index >= 0)
                total += horizontal ? grid.WidthAt(index) : grid.HeightAt(index);
        }

        return total;
    }

    /// <summary>
    /// How much the grid is scaled. A sheet either states a percentage or asks to be squeezed
    /// into so many pages across and down; when it asks, the largest scale that does it is
    /// looked for, never magnifying past what the sheet stated.
    /// </summary>
    private double Scale(SheetPlan plan, GridMeasurer grid)
    {
        if (!_setup.FitToPage)
            return Math.Clamp(_setup.Scale, 10, 400) / 100.0;

        int? wide = _setup.FitToWidth is > 0 ? _setup.FitToWidth : null;
        int? tall = _setup.FitToHeight is > 0 ? _setup.FitToHeight : null;

        if (wide is null && tall is null)
            return 1;

        double scale = 1;
        for (int attempt = 0; attempt < 40; attempt++)
        {
            (List<Band> columns, List<Band> rows) = Bands(plan, grid, scale);
            if (columns.Count <= (wide ?? int.MaxValue) && rows.Count <= (tall ?? int.MaxValue))
                return scale;

            // Ten per cent at a time is what Excel's own dialogue offers, and it stops at a
            // tenth of full size, below which nothing on the page could be read anyway.
            scale -= 0.05;
            if (scale < MinimumScale)
                break;
        }

        _context.Diagnostics.Add(
            PdfExportWarningKind.LayoutApproximated,
            "A sheet could not be squeezed into the pages it asked for and was drawn at a tenth of its size.",
            _sheet.Name);

        return MinimumScale;
    }

    /// <summary>
    /// Cuts a run of lines into bands no taller or wider than the room a page leaves. A line
    /// bigger than a whole page gets a band of its own and runs over the edge, which is what
    /// Excel does with it too.
    /// </summary>
    private static List<Band> Split(int[] lines, Func<int, double> sizeAt, double room, IReadOnlyList<int> breaks)
    {
        room = Math.Max(1, room);
        List<Band> bands = [];
        int from = 0;
        double used = 0;

        for (int i = 0; i < lines.Length; i++)
        {
            double size = sizeAt(i);
            bool forced = i > from && breaks.Contains(lines[i]);

            if (forced || (i > from && used + size > room + 0.01))
            {
                bands.Add(new Band(from, i - from));
                from = i;
                used = 0;
            }

            used += size;
        }

        if (from < lines.Length)
            bands.Add(new Band(from, lines.Length - from));

        return bands.Count == 0 ? [new Band(0, lines.Length)] : bands;
    }

    /// <summary>
    /// The pairings of the bands, in the order the sheet numbers its pages: down the first
    /// column of pages and then across, or across the first row of pages and then down.
    /// </summary>
    private IEnumerable<(Band Columns, Band Rows)> Order(List<Band> columns, List<Band> rows)
    {
        if (_setup.Order == PageOrder.OverThenDown)
        {
            foreach (Band row in rows)
            {
                foreach (Band column in columns)
                    yield return (column, row);
            }

            yield break;
        }

        foreach (Band column in columns)
        {
            foreach (Band row in rows)
                yield return (column, row);
        }
    }
}
