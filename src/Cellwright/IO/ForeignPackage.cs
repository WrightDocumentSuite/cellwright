namespace Cellwright.IO;

/// <summary>
/// Recognises the zip-based documents that are not an xlsx workbook. They all pass the zip
/// signature check and then fail somewhere deeper — an .ods has no content types at all, an
/// .xlsb has a workbook part that is records rather than XML — so without this the reader
/// reports a missing part or a parse error for a file whose actual problem is that it is
/// another format entirely.
/// </summary>
internal static class ForeignPackage
{
    /// <summary>
    /// Says what a workbook part that is not SpreadsheetML actually is, or
    /// <see langword="null"/> when it looks like the XML this reader expects.
    /// </summary>
    public static string? DescribeWorkbookPart(string partPath) =>
        partPath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)
            ? "This is a binary .xlsb workbook: its parts hold records rather than XML, which this reader " +
              "does not parse. Cellwright.Xls does read it — open the file with Spreadsheet.LoadAsync, which " +
              "picks the reader by what the file holds."
            : null;

    /// <summary>
    /// Says what a package holding no workbook part actually is, or <see langword="null"/>
    /// when nothing recognizable explains it.
    /// </summary>
    public static string? Describe(OpcPackage package)
    {
        if (package.PartExists("/mimetype") || package.PartExists("/META-INF/manifest.xml"))
        {
            return "This is an OpenDocument file (.ods and its relatives), not an .xlsx package. The two are " +
                   "both zips and share nothing else. Save it as .xlsx to read it here.";
        }

        if (package.PartExists("/word/document.xml"))
            return "This is a Word document (.docx), not a workbook.";

        if (package.PartExists("/ppt/presentation.xml"))
            return "This is a PowerPoint presentation (.pptx), not a workbook.";

        if (!package.PartExists("/[Content_Types].xml"))
        {
            return "The zip holds no [Content_Types].xml, so it is not an Open Packaging Conventions package " +
                   "at all — just an ordinary archive that happens to start with the same two bytes.";
        }

        return null;
    }
}
