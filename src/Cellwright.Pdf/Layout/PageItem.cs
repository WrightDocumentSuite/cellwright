using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// Something to draw on a page, positioned in the top-down space layout works in.
/// </summary>
/// <remarks>
/// Composition and rendering are kept apart on purpose: composition decides what goes where and
/// may throw a page away and start again, while rendering writes operators and can never be
/// undone. A draw list is the seam between them, and it is deliberately small — a filled
/// rectangle, a stroked segment, a piece of text, a picture, a link.
/// </remarks>
internal abstract class PageItem
{
    /// <summary>The left edge, in points from the left of the page.</summary>
    public double X { get; set; }

    /// <summary>The top edge, in points down from the top of the page.</summary>
    public double Y { get; set; }

    /// <summary>
    /// The structure element this item belongs to, or <see langword="null"/> when it is an
    /// artifact. Only meaningful when the export is tagged.
    /// </summary>
    public TagRef? Tag { get; set; }
}

/// <summary>A solid rectangle: a cell's fill, or the band behind a header.</summary>
internal sealed class FillItem : PageItem
{
    /// <summary>How wide the rectangle is.</summary>
    public required double Width { get; init; }

    /// <summary>How tall the rectangle is.</summary>
    public required double Height { get; init; }

    /// <summary>The colour to fill it with.</summary>
    public required PdfColor Color { get; init; }
}

/// <summary>A straight stroked segment: a grid line or one edge of a cell's border.</summary>
internal sealed class StrokeItem : PageItem
{
    /// <summary>Where the segment ends horizontally, in points from the left of the page.</summary>
    public required double X2 { get; init; }

    /// <summary>Where the segment ends vertically, in points down from the top of the page.</summary>
    public required double Y2 { get; init; }

    /// <summary>How thick the segment is.</summary>
    public required double Thickness { get; init; }

    /// <summary>The colour to stroke it in.</summary>
    public required PdfColor Color { get; init; }

    /// <summary>The line style, which decides the dash pattern and whether the line is doubled.</summary>
    public BorderLineStyle Style { get; init; } = BorderLineStyle.Thin;
}

/// <summary>
/// One line of text drawn at a place already decided. A cell that wraps produces one of these
/// per line; the position is the line's own left edge and the top of its line box.
/// </summary>
internal sealed class TextItem : PageItem
{
    /// <summary>The characters, in the order they are painted.</summary>
    public required string Text { get; init; }

    /// <summary>How to draw them.</summary>
    public required TextAppearance Appearance { get; init; }

    /// <summary>How wide the text is, which is what an underline and a highlight span.</summary>
    public required double Width { get; init; }

    /// <summary>How tall its line box is.</summary>
    public required double Height { get; init; }

    /// <summary>
    /// How the text is turned, in visual clockwise degrees: 0 for ordinary text, 90 for text read
    /// downwards, 270 for text read upwards. For turned text the position is the visual top-left
    /// of the strip it occupies and <see cref="Width"/> runs along the strip.
    /// </summary>
    public int Rotation { get; init; }

    /// <summary>The rectangle beyond which the text is not drawn, or null when it is not clipped.</summary>
    public (double Left, double Top, double Right, double Bottom)? Clip { get; init; }
}

/// <summary>A picture placed on the page.</summary>
internal sealed class ImageItem : PageItem
{
    /// <summary>The picture to draw.</summary>
    public required Picture Picture { get; init; }

    /// <summary>How wide it is drawn.</summary>
    public required double Width { get; init; }

    /// <summary>How tall it is drawn.</summary>
    public required double Height { get; init; }

    /// <summary>
    /// The rectangle beyond which it is not drawn, or null when the whole of it is on the page.
    /// A picture pinned across a page break is drawn on both and cut at the edge of each.
    /// </summary>
    public (double Left, double Top, double Right, double Bottom)? Clip { get; init; }
}

/// <summary>The shapes an icon-set rule draws with.</summary>
/// <remarks>
/// The format names seventeen sets (ST_IconSetType, §18.18.42) but draws them from far fewer
/// shapes: what separates <c>3TrafficLights1</c> from <c>3TrafficLights2</c> is trim, not form.
/// Drawing by shape family covers every named set with the handful of outlines below, and a set
/// this does not know still gets something in the right place rather than nothing.
/// </remarks>
internal enum IconShape
{
    /// <summary>A filled disc: the traffic lights and the red-to-black set.</summary>
    Circle,

    /// <summary>An arrow, turned by <see cref="IconItem.Angle"/>.</summary>
    Arrow,

    /// <summary>A flag on a staff.</summary>
    Flag,

    /// <summary>A triangle, turned by <see cref="IconItem.Angle"/>.</summary>
    Triangle,

    /// <summary>A diamond on its point.</summary>
    Diamond,

    /// <summary>A tick.</summary>
    Check,

    /// <summary>A cross.</summary>
    Cross,

    /// <summary>An exclamation mark.</summary>
    Exclamation,

    /// <summary>Bars of rising height, of which the lower ones are filled.</summary>
    Rating,

    /// <summary>A disc with a fraction of it filled, like a pie.</summary>
    Quarters,

    /// <summary>A five-point star, filled from left to right.</summary>
    Star,

    /// <summary>A short horizontal dash.</summary>
    Dash,

    /// <summary>A two-by-two group of boxes, filled by degree.</summary>
    Boxes,

    /// <summary>An intentional absence of an icon.</summary>
    None,
}

/// <summary>One icon of an icon-set rule, drawn in a cell.</summary>
internal sealed class IconItem : PageItem
{
    /// <summary>Which outline to draw.</summary>
    public required IconShape Shape { get; init; }

    /// <summary>The colour to draw it in.</summary>
    public required PdfColor Color { get; init; }

    /// <summary>The side of the square the icon is drawn inside.</summary>
    public required double Size { get; init; }

    /// <summary>
    /// How far the shape is turned, clockwise from pointing up, for the shapes that point.
    /// </summary>
    public double Angle { get; init; }

    /// <summary>How much of the icon is filled, for the shapes that fill by degree.</summary>
    public double Fraction { get; init; } = 1;

    /// <summary>How many parts the icon divides into, for the shapes that have parts.</summary>
    public int Parts { get; init; } = 1;
}

/// <summary>
/// A filled polygon, in the top-down space layout works in. What charts are drawn from: a pie
/// wedge is a fan of short segments, an area series the region under its line.
/// </summary>
internal sealed class PolygonItem : PageItem
{
    /// <summary>The corners, in order; the path closes itself.</summary>
    public required IReadOnlyList<(double X, double Y)> Points { get; init; }

    /// <summary>The colour to fill with.</summary>
    public required PdfColor Color { get; init; }
}

/// <summary>A clickable area over a cell that carries a hyperlink.</summary>
internal sealed class LinkItem : PageItem
{
    /// <summary>Where the link goes, for a link out of the document.</summary>
    public string? Url { get; init; }

    /// <summary>The target worksheet for an internal workbook link.</summary>
    public Worksheet? TargetSheet { get; init; }

    /// <summary>The target cell for an internal workbook link.</summary>
    public CellRef? TargetCell { get; init; }

    /// <summary>The resolved 0-based page index in the output PDF for an internal link.</summary>
    public int? DestinationPageIndex { get; set; }

    /// <summary>How wide the area is.</summary>
    public required double Width { get; init; }

    /// <summary>How tall it is.</summary>
    public required double Height { get; init; }
}

/// <summary>An interactive comment thread anchored to the visible rectangle of a cell.</summary>
/// <remarks>
/// This item has no content-stream footprint: its icon is a PDF annotation of a fixed viewer size.
/// Keeping it in the page list nevertheless preserves its reading position beside the cell content
/// and lets composition decide the anchor without changing measurement or pagination.
/// </remarks>
internal sealed class CommentItem : PageItem
{
    /// <summary>How wide the visible cell or merge is before the page scale is applied.</summary>
    public required double Width { get; init; }

    /// <summary>How tall the visible cell or merge is before the page scale is applied.</summary>
    public required double Height { get; init; }

    /// <summary>The root and replies to create.</summary>
    public required PdfCommentThread Thread { get; init; }

    /// <summary>Which of several roots sharing this cell this one is.</summary>
    public int StackIndex { get; init; }
}
