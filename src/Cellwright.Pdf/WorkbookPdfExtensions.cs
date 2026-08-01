using Inkwright;

namespace Cellwright.Pdf;

/// <summary>Printing a workbook, from the workbook.</summary>
/// <example>
/// <code>
/// Workbook book = await Workbook.LoadAsync("report.xlsx");
/// PdfExportDiagnostics diagnostics = book.SaveAsPdf("report.pdf");
/// </code>
/// </example>
public static class WorkbookPdfExtensions
{
    extension(Workbook workbook)
    {
        /// <summary>Renders the workbook and writes it to a file.</summary>
        /// <param name="path">Where to write the PDF.</param>
        /// <param name="options">How to render it, or <see langword="null"/> for the defaults.</param>
        public PdfExportDiagnostics SaveAsPdf(string path, PdfExportOptions? options = null) =>
            PdfExporter.Save(workbook, path, options);

        /// <summary>
        /// Renders the workbook and hands back the document, still open, for a caller that has
        /// something else to do to it — sign it, encrypt it, claim a profile — before saving.
        /// </summary>
        /// <param name="options">How to render it, or <see langword="null"/> for the defaults.</param>
        public PdfExportResult ToPdf(PdfExportOptions? options = null) =>
            PdfExporter.Render(workbook, options);
    }

    extension(Worksheet sheet)
    {
        /// <summary>Renders this sheet alone and hands back the document, still open.</summary>
        /// <param name="options">How to render it, or <see langword="null"/> for the defaults.</param>
        public PdfExportResult ToPdf(PdfExportOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            options ??= new PdfExportOptions();
            var one = new PdfExportOptions
            {
                Selection = SheetSelection.Named,
                FallbackFontFamily = options.FallbackFontFamily,
                SubsetFonts = options.SubsetFonts,
                Culture = options.Culture,
                Tagged = options.Tagged,
                Language = options.Language,
                MaxPages = options.MaxPages,
            };

            one.Sheets.Add(sheet.Name);
            foreach ((string family, string file) in options.FontFiles)
                one.FontFiles[family] = file;

            return PdfExporter.Render(sheet.Workbook, one);
        }
    }
}
