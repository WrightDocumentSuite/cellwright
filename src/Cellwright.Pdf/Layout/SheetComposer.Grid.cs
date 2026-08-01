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
            PageReach = _geometry.ContentWidth / (page.Scale > 0 ? page.Scale : 1),
            Conditional = _conditional,
        };
        cells.Fills();

        if (_setup.PrintGridLines)
            GridLines(page, columnSlots, rowSlots);

        cells.Borders();
        cells.Text();
        cells.Links();
        Pictures(page, grid, columnSlots, rowSlots);
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
