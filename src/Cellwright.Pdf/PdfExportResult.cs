using Inkwright;

namespace Cellwright.Pdf;

/// <summary>
/// What one render produced: the document, what had to be worked around, and how much paper it
/// came to.
/// </summary>
/// <remarks>
/// The document is handed over open so that a caller can sign it, encrypt it or claim
/// conformance with a profile before saving. Whoever takes it owns it and has to dispose it.
/// </remarks>
public sealed class PdfExportResult
{
    internal PdfExportResult(PdfDocument document, PdfExportDiagnostics diagnostics, int pageCount)
    {
        Document = document;
        Diagnostics = diagnostics;
        PageCount = pageCount;
    }

    /// <summary>The rendered document, still open.</summary>
    public PdfDocument Document { get; }

    /// <summary>What the exporter had to work around.</summary>
    public PdfExportDiagnostics Diagnostics { get; }

    /// <summary>How many pages it came to.</summary>
    public int PageCount { get; }
}
