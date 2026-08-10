using Cellwright.Model;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// Drawing one page of the grid: the fills behind the cells, the lines between them and the
/// text inside them, in that order, so that nothing is painted over.
/// </summary>
internal sealed partial class SheetComposer
{
    /// <summary>The colour Excel prints its grid lines in.</summary>
    private static readonly PdfColor GridColor = PdfColor.Rgb(0.65f, 0.65f, 0.65f);

    /// <summary>How thick a grid line is.</summary>
    private const double GridThickness = 0.5;

    private void DrawBand(ComposedPage page, SheetPlan plan, GridMeasurer grid, Band columns, Band rows)
    {
        PageSlots columnSlots = Slots(plan.TitleColumns, plan.Columns, columns, grid, horizontal: true);
        PageSlots rowSlots = Slots(plan.TitleRows, plan.Rows, rows, grid, horizontal: false);

        Place(page, columnSlots, rowSlots);

        var cells = new CellPlacer(_context, _sheet, grid, page, columnSlots, rowSlots)
        {
            Tags = _tags,
            Comments = _comments,
            PageReach = _geometry.ContentWidth / (page.Scale > 0 ? page.Scale : 1),
            Conditional = _conditional,
        };

        bool hasConditionalFormats = _conditional.HasRules;
        bool drawStyles = hasConditionalFormats ||
            AnyOccupiedOnPage(_sheet.StyleIndexes, columnSlots, rowSlots);
        bool drawText = hasConditionalFormats || _tags is not null || _comments is not null ||
            AnyOccupiedOnPage(_sheet.Values, columnSlots, rowSlots);

        if (drawStyles)
            cells.Fills();

        if (_setup.PrintGridLines)
            GridLines(page, columnSlots, rowSlots);

        if (drawStyles)
            cells.Borders();

        if (drawText)
            cells.Text();

        cells.Links();
        Pictures(page, grid, columnSlots, rowSlots);
    }

    /// <summary>Whether any occupied grid slot belongs to the band or repeated titles of a page.</summary>
    /// <remarks>
    /// Repeated titles and the page band can sit thousands of rows apart. Treating their common
    /// bounding rectangle as the page would repeatedly scan the gap, so each of the at most four
    /// title/band rectangles is checked separately.
    /// </remarks>
    private static bool AnyOccupiedOnPage<T>(SheetGrid<T> grid, PageSlots columns, PageSlots rows)
    {
        int bandColumns = columns.Count - columns.Titles;
        int bandRows = rows.Count - rows.Titles;

        return Occupied(grid, columns, 0, columns.Titles, rows, 0, rows.Titles) ||
            Occupied(grid, columns, columns.Titles, bandColumns, rows, 0, rows.Titles) ||
            Occupied(grid, columns, 0, columns.Titles, rows, rows.Titles, bandRows) ||
            Occupied(grid, columns, columns.Titles, bandColumns, rows, rows.Titles, bandRows);
    }

    private static bool Occupied<T>(
        SheetGrid<T> grid,
        PageSlots columns,
        int firstColumnSlot,
        int columnCount,
        PageSlots rows,
        int firstRowSlot,
        int rowCount)
    {
        if (columnCount == 0 || rowCount == 0)
            return false;

        (int firstRow, int lastRow) = Bounds(rows.Lines, firstRowSlot, rowCount);
        (int firstColumn, int lastColumn) = Bounds(columns.Lines, firstColumnSlot, columnCount);
        return grid.AnyOccupied(new RangeRef(
            new CellRef(firstRow, firstColumn),
            new CellRef(lastRow, lastColumn)));
    }

    private static (int First, int Last) Bounds(int[] lines, int start, int count)
    {
        int first = lines[start];
        int last = first;
        int end = start + count;

        for (int i = start + 1; i < end; i++)
        {
            first = Math.Min(first, lines[i]);
            last = Math.Max(last, lines[i]);
        }

        return (first, last);
    }

    /// <summary>
    /// The slots of one axis: the repeated titles first, then the band, each with the position
    /// its near edge falls at once the page origin is known.
    /// </summary>
    private static PageSlots Slots(int[] titles, int[] lines, Band band, GridMeasurer grid, bool horizontal)
    {
        List<int> ordered = [];
        foreach (int title in titles)
        {
            // A title inside the band is already there; repeating it would draw it twice.
            int index = horizontal ? grid.ColumnIndex(title) : grid.RowIndex(title);
            if (index >= 0 && (index < band.From || index >= band.From + band.Count))
                ordered.Add(index);
        }

        int repeated = ordered.Count;
        for (int i = band.From; i < band.From + band.Count; i++)
            ordered.Add(i);

        var slots = new PageSlots
        {
            Lines = [.. ordered.Select(index => lines[index])],
            Indices = [.. ordered],
            Near = new double[ordered.Count],
            Size = new double[ordered.Count],
            Titles = repeated,
            First = band.From == 0,
            Last = band.From + band.Count >= lines.Length,
        };

        double at = 0;
        for (int i = 0; i < ordered.Count; i++)
        {
            slots.Near[i] = at;
            slots.Size[i] = horizontal ? grid.WidthAt(ordered[i]) : grid.HeightAt(ordered[i]);
            at += slots.Size[i];
        }

        return slots;
    }

    /// <summary>
    /// Puts the grid on the page: at the top-left of the printed area, or centred in it when the
    /// sheet asks. The offset is applied once, when the page is drawn, rather than to every item.
    /// </summary>
    private void Place(ComposedPage page, PageSlots columns, PageSlots rows)
    {
        double x = _geometry.ContentLeft;
        double y = _geometry.ContentTop;

        if (_setup.CenterHorizontally)
            x += Math.Max(0, (_geometry.ContentWidth - (columns.Extent * page.Scale)) / 2);

        if (_setup.CenterVertically)
            y += Math.Max(0, (_geometry.ContentHeight - (rows.Extent * page.Scale)) / 2);

        page.Offset = (x, y);
    }

    /// <summary>
    /// Draws the grid Excel prints when it is asked to: a line between every pair of printed
    /// cells and around the outside of them, whether or not the cells hold anything.
    /// </summary>
    private static void GridLines(ComposedPage page, PageSlots columns, PageSlots rows)
    {
        if (columns.Count == 0 || rows.Count == 0)
            return;

        double right = columns.Far(columns.Count - 1);
        double bottom = rows.Far(rows.Count - 1);

        for (int i = 0; i <= columns.Count; i++)
        {
            double x = i == columns.Count ? right : columns.Near[i];
            page.Items.Add(Line(x, rows.Near[0], x, bottom));
        }

        for (int i = 0; i <= rows.Count; i++)
        {
            double y = i == rows.Count ? bottom : rows.Near[i];
            page.Items.Add(Line(columns.Near[0], y, right, y));
        }
    }

    private static StrokeItem Line(double x1, double y1, double x2, double y2) => new()
    {
        X = x1,
        Y = y1,
        X2 = x2,
        Y2 = y2,
        Thickness = GridThickness,
        Color = GridColor,
        Style = BorderLineStyle.Hair,
    };
}
