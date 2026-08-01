using Cellwright.Pdf.Layout;
using Cellwright.Pdf.Render;
using Inkwright;

namespace Cellwright.Pdf;

/// <summary>
/// Renders a workbook to PDF the way Excel prints it.
/// </summary>
/// <remarks>
/// <para>
/// Printing a spreadsheet is not printing a document: there is no flow to break, only a grid to
/// cut into pages that fit the paper. The work goes in two passes for each sheet — measure the
/// grid, then walk it — and the two are kept apart so that the second can be repeated at a
/// different scale without measuring anything twice.
/// </para>
/// <para>
/// A render never throws for content it cannot draw. It draws what it can and records the rest
/// in <see cref="PdfExportDiagnostics"/>, so a batch job goes on and a person can see afterwards
/// which files need looking at.
/// </para>
/// </remarks>
public static class PdfExporter
{
    /// <summary>Renders a workbook and hands back the document, still open.</summary>
    /// <param name="workbook">The workbook to render.</param>
    /// <param name="options">How to render it, or <see langword="null"/> for the defaults.</param>
    public static PdfExportResult Render(Workbook workbook, PdfExportOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        options ??= new PdfExportOptions();
        PdfDocument pdf = PdfDocument.Create();

        try
        {
            var context = new PdfExportContext(workbook, pdf, options);
            var composer = new WorkbookComposer(context);
            IReadOnlyList<ComposedPage> pages = composer.Compose();

            DocumentTagger? tagger = options.Tagged ? new DocumentTagger(context) : null;
            new PageRenderer(context).Render(pages, tagger);
            tagger?.Build();

            Metadata.Apply(context);
            return new PdfExportResult(pdf, context.Diagnostics, pages.Count);
        }
        catch
        {
            pdf.Dispose();
            throw;
        }
    }

    /// <summary>Renders a workbook and writes it to a file.</summary>
    /// <param name="workbook">The workbook to render.</param>
    /// <param name="path">Where to write the PDF.</param>
    /// <param name="options">How to render it, or <see langword="null"/> for the defaults.</param>
    public static PdfExportDiagnostics Save(Workbook workbook, string path, PdfExportOptions? options = null)
    {
        PdfExportResult result = Render(workbook, options);
        using PdfDocument document = result.Document;
        document.Save(path);
        return result.Diagnostics;
    }
}
