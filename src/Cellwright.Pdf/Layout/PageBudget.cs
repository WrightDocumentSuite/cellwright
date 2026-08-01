namespace Cellwright.Pdf.Layout;

/// <summary>The ceiling on how many pages one render may produce.</summary>
/// <remarks>
/// The limit is worth nothing if it is only noticed once the pages exist: a sheet that cannot
/// fit its paper would have filled memory with them by then. So it is checked at every point a
/// count is known and before any of that count is built.
/// </remarks>
internal static class PageBudget
{
    /// <summary>Stops a render that has gone past the pages it was allowed.</summary>
    /// <param name="pages">How many pages the render has come to, or will come to.</param>
    /// <param name="options">The options the limit was set in.</param>
    /// <exception cref="InvalidOperationException">The count is past the limit.</exception>
    public static void Check(int pages, PdfExportOptions options)
    {
        if (pages <= options.MaxPages)
            return;

        throw new InvalidOperationException(
            $"The workbook did not finish within {options.MaxPages} pages. A sheet whose " +
            "content cannot fit its paper would otherwise paginate for ever; raise " +
            $"{nameof(PdfExportOptions)}.{nameof(PdfExportOptions.MaxPages)} if the workbook really is this long.");
    }
}
