using Cellwright.Pdf.Layout;
using Cellwright.Styles;
using Inkwright;
using Inkwright.Annotations;
using Inkwright.Content;
using Inkwright.Cos;

namespace Cellwright.Pdf.Render;

/// <summary>Draws one page's items onto its canvas.</summary>
/// <remarks>
/// Everything the layout produced is positioned in the grid's own space, so the scale and the
/// offset of the page are applied once here, as a matrix around the whole body, rather than to
/// every item. The furniture is drawn outside that matrix, at page scale, which is where headers
/// and footers belong unless the sheet says otherwise.
/// </remarks>
internal sealed partial class ItemPainter
{
    /// <summary>How far below the baseline an underline sits, as a fraction of the size.</summary>
    private const double UnderlineOffset = 0.12;

    /// <summary>How thick an underline is, as a fraction of the size.</summary>
    private const double UnderlineWeight = 0.06;

    /// <summary>How far above the baseline a strikethrough sits, as a fraction of the size.</summary>
    private const double StrikeOffset = 0.26;

    private readonly ContentCanvas _canvas;
    private readonly PdfPage _page;
    private readonly PageGeometry _geometry;
    private readonly ImageEmbedder _images;
    private readonly ITagSink? _tags;

    private double _scale = 1;
    private (double X, double Y) _offset;

    internal ItemPainter(
        ContentCanvas canvas, PdfPage page, ComposedPage composed, ImageEmbedder images, ITagSink? tags)
    {
        _canvas = canvas;
        _page = page;
        _geometry = composed.Geometry;
        _images = images;
        _tags = tags;
    }

    /// <summary>
    /// Starts drawing the body, which the renderer scales and shifts with one matrix around the
    /// whole of it.
    /// </summary>
    /// <param name="scale">How much the body is scaled.</param>
    /// <param name="offset">Where the scaled body starts on the paper.</param>
    /// <remarks>
    /// That matrix belongs to the content stream, and an annotation is not in the content stream:
    /// its rectangle is read in page space whatever the stream is doing. So anything drawn as an
    /// annotation has to be moved here, by hand, to where the matrix would have put it.
    /// </remarks>
    public void EnterBody(double scale, (double X, double Y) offset)
    {
        _scale = scale is > 0 ? scale : 1;
        _offset = offset;
    }

    /// <summary>Draws one item.</summary>
    /// <param name="item">The item to draw.</param>
    /// <remarks>
    /// In a tagged document every mark on the page must either belong to the structure tree or
    /// say that it belongs to nobody. Fills, grid lines and borders are the latter, so they are
    /// wrapped in an artifact sequence rather than left unaccounted for.
    /// </remarks>
    public void Paint(PageItem item)
    {
        ArgumentNullException.ThrowIfNull(item);

        bool artifact = _tags is not null && item.Tag is null && item is not (LinkItem or CommentItem);
        if (artifact)
            _canvas.BeginMarkedContent(PdfName.Get("Artifact"));

        Draw(item);

        if (artifact)
            _canvas.EndMarkedContent();
    }

    private void Draw(PageItem item)
    {
        switch (item)
        {
            case FillItem fill:
                PaintFill(fill);
                break;
            case StrokeItem stroke:
                PaintStroke(stroke);
                break;
            case TextItem text:
                PaintText(text);
                break;
            case ImageItem image:
                PaintImage(image);
                break;
            case IconItem icon:
                PaintIcon(icon);
                break;
            case PolygonItem polygon:
                PaintPolygon(polygon);
                break;
            case LinkItem link:
                PaintLink(link);
                break;
            case CommentItem comment:
                PaintComment(comment);
                break;
            default:
                break;
        }
    }

    private void PaintFill(FillItem fill)
    {
        if (fill.Width <= 0 || fill.Height <= 0)
            return;

        _canvas.Save()
               .FillColor(fill.Color)
               .Rectangle(fill.X, _geometry.ToPdfY(fill.Y + fill.Height), fill.Width, fill.Height)
               .Fill()
               .Restore();
    }

    private void PaintStroke(StrokeItem stroke)
    {
        if (stroke.Thickness <= 0)
            return;

        _canvas.Save().StrokeColor(stroke.Color).LineWidth(stroke.Thickness);

        if (Dash(stroke.Style, stroke.Thickness) is { Length: > 0 } dash)
            _canvas.Dash(0, dash);

        double y1 = _geometry.ToPdfY(stroke.Y);
        double y2 = _geometry.ToPdfY(stroke.Y2);

        if (stroke.Style == BorderLineStyle.Double)
        {
            // Two thin lines a point apart, straddling the place the single line would go.
            (double dx, double dy) = Normal(stroke.X, y1, stroke.X2, y2, 1);
            _canvas.LineWidth(0.4);
            Segment(stroke.X + dx, y1 + dy, stroke.X2 + dx, y2 + dy);
            Segment(stroke.X - dx, y1 - dy, stroke.X2 - dx, y2 - dy);
        }
        else
        {
            Segment(stroke.X, y1, stroke.X2, y2);
        }

        _canvas.Restore();
    }

    private void Segment(double x1, double y1, double x2, double y2) =>
        _canvas.MoveTo(x1, y1).LineTo(x2, y2).Stroke();

    /// <summary>The offset that separates the two halves of a doubled line.</summary>
    private static (double Dx, double Dy) Normal(double x1, double y1, double x2, double y2, double distance)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt((dx * dx) + (dy * dy));
        return length < 1e-6 ? (0, 0) : (-dy / length * distance / 2, dx / length * distance / 2);
    }

    /// <summary>The dash pattern a border style draws with, or nothing for a solid line.</summary>
    private static double[] Dash(BorderLineStyle style, double thickness) => style switch
    {
        BorderLineStyle.Dotted => [thickness, thickness * 2],
        BorderLineStyle.Dashed or BorderLineStyle.MediumDashed => [thickness * 3, thickness * 2],
        BorderLineStyle.DashDot or BorderLineStyle.MediumDashDot or BorderLineStyle.SlantDashDot =>
            [thickness * 3, thickness * 2, thickness, thickness * 2],
        BorderLineStyle.DashDotDot or BorderLineStyle.MediumDashDotDot =>
            [thickness * 3, thickness * 2, thickness, thickness * 2, thickness, thickness * 2],
        _ => [],
    };

    private void PaintImage(ImageItem item)
    {
        if (item.Width <= 0 || item.Height <= 0 || _images.Embed(item.Picture) is not { } image)
            return;

        int mcid = BeginTag(item.Tag);

        bool clipped = item.Clip is not null;
        if (clipped)
            Clip(item.Clip!.Value);

        _canvas.DrawImage(image, PdfRectangle.FromSize(
            item.X, _geometry.ToPdfY(item.Y + item.Height), item.Width, item.Height));

        if (clipped)
            _canvas.Restore();

        EndTag(item.Tag, mcid);
    }

    private void PaintPolygon(PolygonItem polygon)
    {
        if (polygon.Points.Count < 3)
            return;

        _canvas.Save().FillColor(polygon.Color);
        _canvas.MoveTo(polygon.Points[0].X, _geometry.ToPdfY(polygon.Points[0].Y));
        for (int i = 1; i < polygon.Points.Count; i++)
            _canvas.LineTo(polygon.Points[i].X, _geometry.ToPdfY(polygon.Points[i].Y));

        _canvas.ClosePath().Fill().Restore();
    }

    private void PaintLink(LinkItem link)
    {
        if (string.IsNullOrEmpty(link.Url) || link.Width <= 0 || link.Height <= 0)
            return;

        double x = _offset.X + (link.X * _scale);
        double y = _offset.Y + (link.Y * _scale);
        double width = link.Width * _scale;
        double height = link.Height * _scale;

        _page.Annotations.AddLink(
            PdfRectangle.FromSize(x, _geometry.ToPdfY(y + height), width, height), link.Url);
    }

    private void PaintComment(CommentItem item)
    {
        if (item.Width <= 0 || item.Height <= 0 || item.Thread.Messages.Count == 0)
            return;

        // Centre the compact marker on the real top-right endpoint: only its lower-left quarter
        // overlaps the cell, preserving the anchor without covering a narrow cell's value. The
        // crop clamp keeps edge cells entirely on the physical page.
        const double NoteSize = 14;
        double right = _offset.X + ((item.X + item.Width) * _scale);
        double stagger = Math.Min(item.StackIndex, 4) * 3;
        double down = Math.Min(stagger, Math.Max(0, (item.Height * _scale) - 1));
        PdfRectangle crop = _page.CropBox;
        double x = Math.Clamp(right - (NoteSize / 2), crop.Left, crop.Right - NoteSize);
        double endpointY = _geometry.ToPdfY(_offset.Y + (item.Y * _scale) + down);
        double y = Math.Clamp(
            endpointY - (NoteSize / 2),
            crop.Bottom,
            crop.Top - NoteSize);
        PdfRectangle marker = PdfRectangle.FromSize(x, y, NoteSize, NoteSize);

        var annotations = new PdfMarkupAnnotation[item.Thread.Messages.Count];
        for (int i = 0; i < item.Thread.Messages.Count; i++)
        {
            PdfCommentMessage message = item.Thread.Messages[i];
            PdfMarkupAnnotation annotation;

            if (message.ParentIndex < 0)
            {
                annotation = _page.Annotations.AddNote(
                    new PdfPoint(x, marker.Top),
                    message.Text,
                    message.Author,
                    PdfNoteIcon.Comment);
                annotation.Rectangle = marker;
            }
            else
            {
                if (message.ParentIndex >= i || annotations[message.ParentIndex] is not { } parent)
                    continue;

                annotation = parent.Reply(message.Text, message.Author);
            }

            Describe(annotation, message);
            annotations[i] = annotation;
            _tags?.AddAnnotation(item.Tag, annotation);
        }
    }

    private static void Describe(PdfMarkupAnnotation annotation, PdfCommentMessage message)
    {
        // AddNote/Reply provide a current timestamp and Reply can inherit an author's name. Those
        // are authoring defaults, not source provenance, so overwrite them even when it means null.
        annotation.Author = message.Author;
        annotation.Name = message.Name;
        annotation.Subject = message.Subject;
        // Author and Subject deliberately stamp /M as authoring APIs. Restore the source dates
        // last so that a missing date stays missing and a known one is not replaced with now.
        annotation.CreatedAt = message.Created;
        annotation.ModifiedAt = message.Created;
    }

    private int BeginTag(TagRef? tag)
    {
        if (tag is null || _tags is null)
            return -1;

        int mcid = _tags.Next(tag);
        var properties = new PdfDictionary(1);
        properties.Set(PdfName.Get("MCID"), PdfValue.Integer(mcid));
        _canvas.BeginMarkedContent(PdfName.Get(tag.Tag), properties);
        return mcid;
    }

    private void EndTag(TagRef? tag, int mcid)
    {
        if (tag is not null && _tags is not null && mcid >= 0)
            _canvas.EndMarkedContent();
    }
}

/// <summary>
/// Where the painter reports the marked-content sequences it writes, so a structure tree can
/// point at them. Only used when the export is tagged.
/// </summary>
internal interface ITagSink
{
    /// <summary>Reserves the next marked-content identifier and attaches it to a structure element.</summary>
    /// <param name="tag">The element the sequence belongs to.</param>
    int Next(TagRef tag);

    /// <summary>Places an annotation in reading order under its semantic owner.</summary>
    void AddAnnotation(TagRef? tag, PdfAnnotation annotation);
}
