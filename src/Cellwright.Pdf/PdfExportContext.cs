using System.Globalization;
using Cellwright.Pdf.Fonts;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf;

/// <summary>
/// Everything one render needs and everything it accumulates: the workbook, the document being
/// built, the fonts already embedded, and the compromises made along the way.
/// </summary>
internal sealed class PdfExportContext : IDisposable
{
    /// <summary>The scratch document a measuring context owns, or null when it renders into someone else's.</summary>
    private readonly PdfDocument? _scratch;

    internal PdfExportContext(Workbook source, PdfDocument pdf, PdfExportOptions options)
        : this(source, pdf, options, scratch: null)
    {
    }

    private PdfExportContext(Workbook source, PdfDocument pdf, PdfExportOptions options, PdfDocument? scratch)
    {
        Source = source;
        Pdf = pdf;
        Options = options;
        _scratch = scratch;
        Fonts = new FontMap(pdf, options, Diagnostics);
    }

    /// <summary>
    /// A context that only measures: it answers the same font metrics a render would use, without
    /// a document anybody saves.
    /// </summary>
    /// <remarks>
    /// Inkwright binds a <see cref="Inkwright.Fonts.PdfFont"/> to the document it was embedded into, so there is no
    /// such thing as a face without one — asking for metrics means embedding somewhere. The scratch
    /// document is that somewhere, and it is thrown away with the context. The point of going
    /// through the ordinary path rather than a metrics-only shortcut is that measuring and printing
    /// then cannot disagree: an auto-fitted column is exactly as wide as the column that gets
    /// printed, substituted faces and all.
    /// </remarks>
    /// <param name="source">The workbook being measured.</param>
    /// <param name="options">The options a render of it would use.</param>
    public static PdfExportContext ForMeasurement(Workbook source, PdfExportOptions options)
    {
        PdfDocument scratch = PdfDocument.Create();
        return new PdfExportContext(source, scratch, options, scratch);
    }

    /// <summary>The workbook being rendered.</summary>
    public Workbook Source { get; }

    /// <summary>The document being built.</summary>
    public PdfDocument Pdf { get; }

    /// <summary>How to render it.</summary>
    public PdfExportOptions Options { get; }

    /// <summary>The fonts, embedded once each.</summary>
    public FontMap Fonts { get; }

    /// <summary>What had to be worked around.</summary>
    public PdfExportDiagnostics Diagnostics { get; } = new();

    /// <summary>The culture number formats are rendered in.</summary>
    public CultureInfo Culture => Options.Culture;

    /// <summary>
    /// A workbook colour as the page draws it: a theme slot resolved through the workbook's own
    /// theme and tinted, an indexed colour through the legacy palette, and the automatic colour
    /// left to whatever the caller says it means in this position.
    /// </summary>
    /// <param name="color">The colour a style names.</param>
    /// <param name="fallback">What the automatic colour means here.</param>
    public PdfColor ColorOf(ExcelColor color, PdfColor fallback)
    {
        if (color.IsAuto)
            return fallback;

        ExcelColor resolved = color.IsThemed ? Source.Theme.Resolve(color) : color;
        return resolved.IsAuto ? fallback : PdfColor.FromRgb(resolved.Argb & 0x00FFFFFFu);
    }

    /// <summary>Whether a colour would draw as anything at all.</summary>
    /// <param name="color">The colour a style names.</param>
    public static bool HasColor(ExcelColor color) => !color.IsAuto;

    /// <summary>Releases the scratch document, if this context made one. Rendering contexts hold none.</summary>
    public void Dispose() => _scratch?.Dispose();
}
