using System.Globalization;
using System.Xml;
using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Xml;

namespace Cellwright.Model;

/// <summary>
/// Reads the external link parts (§18.14) into <see cref="ExternalBook"/> models: the sheet
/// names and the values Excel cached at the last refresh. The parts themselves still travel
/// through preservation untouched; this is the readable view formulas answer from.
/// </summary>
internal static class ExternalBookLoader
{
    public static async ValueTask<IReadOnlyList<ExternalBook>> LoadAsync(
        XlsxReader reader, Action<WorkbookWarning> warn, CancellationToken cancellationToken)
    {
        IReadOnlyList<string> order = await ReadReferenceOrderAsync(reader, cancellationToken).ConfigureAwait(false);
        if (order.Count == 0)
            return [];

        Dictionary<string, OpcRelationship> byId = [];
        foreach (OpcRelationship relationship in
            await reader.Package.GetRelationshipsAsync(reader.WorkbookPartPath, cancellationToken).ConfigureAwait(false))
        {
            byId[relationship.Id] = relationship;
        }

        List<ExternalBook> books = [];
        foreach (string relationshipId in order)
        {
            ExternalBook? book = null;
            if (byId.TryGetValue(relationshipId, out OpcRelationship relationship) && !relationship.IsExternal)
            {
                string partPath = OpcPath.Resolve(reader.WorkbookPartPath, relationship.Target);
                try
                {
                    book = await LoadBookAsync(reader, partPath, cancellationToken).ConfigureAwait(false);
                }
                catch (XmlException exception)
                {
                    warn(new WorkbookWarning(
                        "BadExternalLink",
                        $"The external link part '{partPath}' could not be read: {exception.Message}"));
                }
            }

            // The position in this list is the [N] index formulas use, so a link that failed
            // to load still occupies its slot.
            books.Add(book ?? new ExternalBook(null, [], []));
        }

        return books;
    }

    /// <summary>
    /// The order of <c>externalReference</c> elements in workbook.xml — the numbering behind
    /// <c>[1]Sheet1!A1</c>. The workbook part is small and buffered, so it is simply re-read.
    /// </summary>
    private static async ValueTask<IReadOnlyList<string>> ReadReferenceOrderAsync(
        XlsxReader reader, CancellationToken cancellationToken)
    {
        byte[] content = await reader.ReadPartBytesAsync(reader.WorkbookPartPath, cancellationToken).ConfigureAwait(false);
        List<string> order = [];
        using var stream = new MemoryStream(content, writable: false);
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "externalReference")
                continue;

            string? id = xml.GetAttribute("r:id") ?? xml.GetAttribute(
                "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (id is not null)
                order.Add(id);
        }

        return order;
    }

    private static async ValueTask<ExternalBook?> LoadBookAsync(
        XlsxReader reader, string partPath, CancellationToken cancellationToken)
    {
        if (!reader.PartExists(partPath))
            return null;

        // The target path of the external book hangs off the link part's own relationships.
        string? targetPath = null;
        foreach (OpcRelationship relationship in
            await reader.Package.GetRelationshipsAsync(partPath, cancellationToken).ConfigureAwait(false))
        {
            if (relationship.Type.EndsWith("/externalLinkPath", StringComparison.Ordinal))
                targetPath = relationship.Target;
        }

        byte[] content = await reader.ReadPartBytesAsync(partPath, cancellationToken).ConfigureAwait(false);
        using var stream = new MemoryStream(content, writable: false);
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);

        List<string> sheetNames = [];
        List<ExternalSheetCache> sheets = [];
        ExternalSheetCache? current = null;
        int currentRow = 0;
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "sheetName" when xml.GetAttribute("val") is { } name:
                    sheetNames.Add(name);
                    sheets.Add(new ExternalSheetCache());
                    break;

                case "sheetData":
                    int sheetId = ParseInt(xml.GetAttribute("sheetId"));
                    current = (uint)sheetId < (uint)sheets.Count ? sheets[sheetId] : null;
                    break;

                case "row":
                    currentRow = ParseInt(xml.GetAttribute("r"));
                    break;

                case "cell":
                    ReadCell(xml, current, currentRow);
                    break;

                default:
                    break;
            }
        }

        return new ExternalBook(targetPath, sheetNames, sheets);
    }

    private static void ReadCell(XmlReader xml, ExternalSheetCache? sheet, int row)
    {
        if (sheet is null)
            return;

        string? reference = xml.GetAttribute("r");
        string? type = xml.GetAttribute("t");
        if (!CellRef.TryParse(reference, null, out CellRef cell))
            return;

        string text = ReadValueText(xml);
        CellValue value = type switch
        {
            "b" => CellValue.Boolean(text == "1"),
            "e" => ExcelErrorText.TryParse(text, out ExcelError error) ? CellValue.Error(error) : CellValue.Blank,
            "str" => CellValue.Text(OoxmlText.Decode(text)),
            _ => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                ? CellValue.Number(number)
                : CellValue.Blank,
        };

        int cellRow = cell.Row > 0 ? cell.Row : row;
        if (cellRow > 0 && cell.Column > 0 && !value.IsBlank)
            sheet.Add(cellRow, cell.Column, value);
    }

    private static string ReadValueText(XmlReader xml)
    {
        if (xml.IsEmptyElement)
            return "";

        using XmlReader subtree = xml.ReadSubtree();
        subtree.Read();
        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "v")
                return subtree.ReadElementContentAsString();
        }

        return "";
    }

    private static int ParseInt(string? text) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) ? value : 0;
}
