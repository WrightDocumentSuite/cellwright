using Inkwright.Images;

namespace Cellwright.Pdf.Render;

/// <summary>
/// Puts the pictures a workbook uses into the PDF, once each however often they are drawn.
/// </summary>
/// <remarks>
/// Inkwright embeds JPEG and PNG without re-encoding them, which is the whole point: a photograph
/// keeps its compression and a screenshot keeps its palette. Anything else — a metafile above
/// all, which is what a chart pasted into a sheet often is — would have to be decoded and drawn
/// again; this says so in the diagnostics and leaves the space blank rather than drawing
/// something wrong.
/// </remarks>
internal sealed class ImageEmbedder
{
    private readonly PdfExportContext _context;
    private readonly Dictionary<Picture, PdfImage?> _cache = [];

    internal ImageEmbedder(PdfExportContext context) => _context = context;

    /// <summary>The embedded picture, or <see langword="null"/> when the format cannot be embedded.</summary>
    /// <param name="picture">The picture the workbook carries.</param>
    public PdfImage? Embed(Picture picture)
    {
        ArgumentNullException.ThrowIfNull(picture);

        if (_cache.TryGetValue(picture, out PdfImage? cached))
            return cached;

        PdfImage? embedded = TryEmbed(picture);
        _cache[picture] = embedded;
        return embedded;
    }

    private PdfImage? TryEmbed(Picture picture)
    {
        try
        {
            return PdfImage.Load(_context.Pdf, picture.Content.Span);
        }
        catch (NotSupportedException)
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.ImageSkipped,
                $"Pictures in {picture.Format} are not embedded; only JPEG and PNG are.",
                picture.Format.ToString());
            return null;
        }
        catch (Exception exception) when (exception is InvalidDataException or ArgumentException)
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.ImageSkipped,
                $"A picture in {picture.Format} could not be read: {exception.Message}",
                picture.Name ?? picture.Format.ToString());
            return null;
        }
    }
}
