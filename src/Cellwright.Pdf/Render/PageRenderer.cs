using Cellwright.Pdf.Layout;
using Inkwright;
using Inkwright.Content;

namespace Cellwright.Pdf.Render;

/// <summary>
/// Writes composed pages into the document: one PDF page per composed page, one content stream
/// per page, and the annotations that cannot live in a content stream.
/// </summary>
internal sealed class PageRenderer
{
    private readonly PdfExportContext _context;
    private readonly ImageEmbedder _images;

    internal PageRenderer(PdfExportContext context)
    {
        _context = context;
        _images = new ImageEmbedder(context);
    }

    /// <summary>Draws every composed page.</summary>
    /// <param name="pages">The pages, in order.</param>
    /// <param name="tags">Where to report marked content, or <see langword="null"/> when untagged.</param>
    public void Render(IReadOnlyList<ComposedPage> pages, ITagWriter? tags = null)
    {
        ArgumentNullException.ThrowIfNull(pages);

        var pdfPages = new PdfPage[pages.Count];
        for (int i = 0; i < pages.Count; i++)
        {
            ComposedPage composed = pages[i];
            pdfPages[i] = _context.Pdf.Pages.Add(new PageSize(composed.Geometry.Width, composed.Geometry.Height));
        }

        for (int i = 0; i < pages.Count; i++)
            RenderPage(pages[i], pdfPages[i], tags);
    }

    private void RenderPage(ComposedPage composed, PdfPage page, ITagWriter? tags)
    {
        ITagSink? sink = tags?.ForPage(page, composed);

        using var canvas = ContentCanvas.ForPage(page);
        var painter = new ItemPainter(canvas, page, composed, _images, sink);

        // The furniture is drawn at page scale; the body carries the sheet's own scale and
        // wherever the grid was placed on the paper.
        foreach (PageItem item in composed.Furniture)
            painter.Paint(item);

        canvas.Save();
        canvas.Transform(PdfMatrix.Translation(composed.Offset.X, -composed.Offset.Y));

        if (composed.Scale is > 0 and not 1)
            canvas.Transform(ScaleAbout(composed.Scale, composed.Geometry.Height));

        painter.EnterBody(composed.Scale, composed.Offset);

        foreach (PageItem item in composed.Items)
            painter.Paint(item);

        canvas.Restore();
        canvas.Commit();
    }

    /// <summary>
    /// The scale a page prints its grid at, about the top-left corner of the paper. Layout runs
    /// top-down and PDF bottom-up, so scaling about the origin would move the grid off the top
    /// of the page; anchoring it at the top is what a printer does.
    /// </summary>
    private static PdfMatrix ScaleAbout(double scale, double height) =>
        new(scale, 0, 0, scale, 0, height * (1 - scale));
}

/// <summary>Builds the structure tree as pages are drawn. Implemented only when tagging is on.</summary>
internal interface ITagWriter
{
    /// <summary>Starts collecting the marked content of one page.</summary>
    /// <param name="page">The page being drawn.</param>
    /// <param name="composed">What was composed for it.</param>
    public ITagSink ForPage(PdfPage page, ComposedPage composed);
}
