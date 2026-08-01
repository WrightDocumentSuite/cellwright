namespace Cellwright.Pdf.Layout;

/// <summary>
/// The pictures floating over a sheet's grid, placed by the cells they are pinned to.
/// </summary>
/// <remarks>
/// A picture pinned between two cells is as big as those cells make it, so widening a column
/// stretches it; one pinned to a single cell keeps the size it was given. Either way the anchor
/// counts from a cell's top-left corner in English Metric Units, which is how a spreadsheet
/// measures anything that is not a cell.
/// </remarks>
internal sealed partial class SheetComposer
{
    private void Pictures(ComposedPage page, GridMeasurer grid, PageSlots columns, PageSlots rows)
    {
        if (_sheet.Pictures.Count == 0 && _sheet.Charts.Count == 0)
            return;

        var across = Axis.Of(columns, grid.ColumnPosition, _geometry.ContentWidth / Scaled(page));
        var down = Axis.Of(rows, grid.RowPosition, _geometry.ContentHeight / Scaled(page));

        foreach (Picture picture in _sheet.Pictures)
        {
            if (picture.IsHidden)
                continue;

            if (Frame(picture.Anchor, across, down) is not { } frame)
                continue;

            page.Items.Add(new ImageItem
            {
                Picture = picture,
                X = frame.X,
                Y = frame.Y,
                Width = frame.Width,
                Height = frame.Height,
                Clip = frame.Clip,
                Tag = _tags is null ? null : new TagRef("Figure", picture) { AlternateText = picture.Description },
            });
        }

        Charts(page, across, down);
    }

    /// <summary>
    /// The charts of the sheet, drawn where their anchors put them. A chart is drawn on the
    /// page that holds the whole of it; one pinned across a page break has no half worth
    /// showing — half a plot reads as the plot — and is reported instead.
    /// </summary>
    private void Charts(ComposedPage page, Axis across, Axis down)
    {
        if (_sheet.Charts.Count == 0)
            return;

        Charts.ChartRenderer? renderer = null;
        foreach (Chart chart in _sheet.Charts)
        {
            if (Frame(chart.Anchor, across, down) is not { } frame)
                continue;

            if (frame.Clip is not null)
            {
                _context.Diagnostics.Add(
                    PdfExportWarningKind.LayoutApproximated,
                    "A chart pinned across a page break is not drawn; move it or widen the page.",
                    "chart-split");
                continue;
            }

            if (chart.Kind == ChartKind.Unknown || chart.Series.Count == 0)
            {
                _context.Diagnostics.Add(
                    PdfExportWarningKind.ContentSkipped,
                    "A chart of a kind this renderer does not draw was left out.",
                    $"chart-{chart.Kind}");
                continue;
            }

            renderer ??= new Charts.ChartRenderer(_context, _appearances);
            renderer.Draw(page, chart, frame.X, frame.Y, frame.Width, frame.Height);
        }
    }

    private static double Scaled(ComposedPage page) => page.Scale > 0 ? page.Scale : 1;

    /// <summary>Where a picture is drawn on one page, and how much of it that page shows.</summary>
    /// <param name="X">Its left edge, in the page's own space.</param>
    /// <param name="Y">Its top edge.</param>
    /// <param name="Width">How wide it is drawn.</param>
    /// <param name="Height">How tall.</param>
    /// <param name="Clip">What of it the page may show, or null when all of it may be.</param>
    private readonly record struct Placement(
        double X,
        double Y,
        double Width,
        double Height,
        (double Left, double Top, double Right, double Bottom)? Clip);

    /// <summary>
    /// Where a picture sits on this page and how much of it is shown, or <see langword="null"/>
    /// when none of it falls here.
    /// </summary>
    /// <remarks>
    /// A picture is placed whether or not the cells it hangs from are on this page: one pinned
    /// across a page break belongs to both pages, and each shows the part of it that lands on
    /// the paper. What the page cannot hold is cut off rather than dropped, which is what a
    /// printer does with it.
    /// </remarks>
    private Placement? Frame(PictureAnchor anchor, Axis across, Axis down)
    {
        (double x, double y, double width, double height) = Placed(anchor, across, down);
        if (width <= 0 || height <= 0)
            return null;

        if (!across.Holds(x, x + width) || !down.Holds(y, y + height))
            return null;

        bool whole = across.Wraps(x, x + width) && down.Wraps(y, y + height);
        return new Placement(
            x,
            y,
            width,
            height,
            whole ? null : (across.Near, down.Near, across.Far, down.Far));
    }

    /// <summary>Where a picture falls in this page's own space, on the paper or off it.</summary>
    private static (double X, double Y, double Width, double Height) Placed(
        PictureAnchor anchor, Axis across, Axis down)
    {
        if (anchor.Kind == PictureAnchorKind.Absolute)
        {
            // An absolute anchor measures from the top-left of the sheet, which is where the
            // grid begins.
            return (
                across.At(Points(anchor.Position.X)),
                down.At(Points(anchor.Position.Y)),
                Points(anchor.Extent.Width),
                Points(anchor.Extent.Height));
        }

        double left = across.AtLine(anchor.From.Column) + Points(anchor.FromOffset.X);
        double top = down.AtLine(anchor.From.Row) + Points(anchor.FromOffset.Y);

        if (anchor.Kind == PictureAnchorKind.OneCell)
            return (left, top, Points(anchor.Extent.Width), Points(anchor.Extent.Height));

        double right = across.AtLine(anchor.To.Column) + Points(anchor.ToOffset.X);
        double bottom = down.AtLine(anchor.To.Row) + Points(anchor.ToOffset.Y);

        return (left, top, right - left, bottom - top);
    }

    private static double Points(long emu) => emu / SheetUnits.EmuPerPoint;

    /// <summary>
    /// How one axis of the sheet maps onto one axis of this page.
    /// </summary>
    /// <param name="Slots">The lines this page draws along the axis.</param>
    /// <param name="Position">How far along the grid a line of the sheet starts.</param>
    /// <param name="PageOrigin">Where the band this page shows begins on the paper.</param>
    /// <param name="GridOrigin">Where that same place is in the grid's own reckoning.</param>
    /// <param name="Near">Where what this page may show begins.</param>
    /// <param name="Far">Where it ends.</param>
    private readonly record struct Axis(
        PageSlots Slots,
        Func<int, double> Position,
        double PageOrigin,
        double GridOrigin,
        double Near,
        double Far)
    {
        /// <summary>
        /// How the axis of a page relates to the axis of the sheet, and how much of the sheet the
        /// page is allowed to show.
        /// </summary>
        /// <param name="slots">The lines this page draws.</param>
        /// <param name="position">How far along the grid a line of the sheet starts.</param>
        /// <param name="room">How much of the paper is printable, in the grid's own measure.</param>
        /// <remarks>
        /// A page shows its own band and no more: a picture hanging past the break is cut there,
        /// and the page after it carries the rest. The first and last bands are the exception —
        /// beyond them there is no other page to carry anything, so they run to the edge of the
        /// paper, which is where a picture wider than the columns it hangs over ends up.
        /// </remarks>
        public static Axis Of(PageSlots slots, Func<int, double> position, double room)
        {
            if (slots.Count == 0)
                return new Axis(slots, position, 0, 0, 0, room);

            int band = Math.Min(slots.Titles, slots.Count - 1);
            double origin = slots.Near[band];

            return new Axis(
                slots,
                position,
                origin,
                position(slots.Lines[band]),
                slots.First ? 0 : origin,
                slots.Last ? room : Math.Min(room, slots.Far(slots.Count - 1)));
        }

        /// <summary>Where a place in the grid falls on this page.</summary>
        public double At(double position) => PageOrigin + (position - GridOrigin);

        /// <summary>
        /// Where a line of the sheet falls on this page: at its own slot when the page draws it,
        /// which is what keeps a picture in a repeated title where the title is, and against the
        /// band otherwise.
        /// </summary>
        public double AtLine(int line)
        {
            int slot = Slots.SlotOf(line);
            return slot >= 0 ? Slots.Near[slot] : At(Position(line));
        }

        /// <summary>Whether any of a run falls on the paper.</summary>
        public bool Holds(double from, double to) => to > Near && from < Far;

        /// <summary>Whether the whole of a run does.</summary>
        public bool Wraps(double from, double to) => from >= Near - 0.01 && to <= Far + 0.01;
    }
}
