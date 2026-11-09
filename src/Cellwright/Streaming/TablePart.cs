using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing a table part. The model keeps the shape of a table (range, columns,
/// banding); filter state, sort state and calculated-column formulas are the file's business
/// and ride along in its own part for as long as the model leaves the table alone.
/// </summary>
internal static class TablePart
{
    /// <summary>Parses a table part, or returns <see langword="null"/> when it declares no usable range.</summary>
    public static ExcelTable? Read(Stream stream)
    {
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        List<string> columns = [];
        string? name = null;
        string? displayName = null;
        string? styleName = null;
        RangeRef range = default;
        bool hasRange = false;
        bool headerRow = true;
        bool totalsRow = false;
        bool rowStripes = false;
        bool columnStripes = false;
        bool filtered = false;
        int id = 0;

        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                // A filterColumn is the table's own filter at work, and the rows it rejects
                // are hidden with nothing else to say so.
                case "filterColumn":
                    filtered = true;
                    break;

                case "table":
                    name = xml.GetAttribute("name") ?? xml.GetAttribute("displayName");
                    displayName = xml.GetAttribute("displayName");
                    hasRange = RangeRef.TryParse(xml.GetAttribute("ref"), null, out range);
                    headerRow = xml.GetAttribute("headerRowCount") is not "0";
                    totalsRow = int.TryParse(xml.GetAttribute("totalsRowCount"), NumberStyles.None, CultureInfo.InvariantCulture, out int totals)
                        && totals > 0;
                    id = int.TryParse(xml.GetAttribute("id"), NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
                    break;

                // Column names travel through the escape convention, so a header holding a line
                // break arrives as _x000a_ and would never match a structured reference.
                case "tableColumn":
                    columns.Add(xml.GetAttribute("name") is { } columnName
                        ? OoxmlText.Decode(columnName)
                        : $"Column{columns.Count + 1}");
                    break;

                case "tableStyleInfo":
                    styleName = xml.GetAttribute("name");
                    rowStripes = xml.GetAttribute("showRowStripes") is "1" or "true";
                    columnStripes = xml.GetAttribute("showColumnStripes") is "1" or "true";
                    break;

                default:
                    break;
            }
        }

        if (name is null || !hasRange || columns.Count == 0)
            return null;

        return new ExcelTable
        {
            Name = name,
            Range = range,
            Columns = columns,
            HasHeaderRow = headerRow,
            HasTotalsRow = totalsRow,
            StyleName = styleName,
            ShowRowStripes = rowStripes,
            ShowColumnStripes = columnStripes,
            DisplayName = displayName,
            Id = id,
            FilterIsActive = filtered,
        };
    }

    /// <summary>Writes a table part.</summary>
    public static void Write(Utf8XmlWriter xml, ExcelTable table)
    {
        // Excel repairs a table whose ref and tableColumns disagree, dropping the table.
        if (table.Columns.Count != table.Range.ColumnCount)
        {
            throw new InvalidOperationException(
                $"Table '{table.Name}' declares {table.Columns.Count} column(s) over a range {table.Range.ColumnCount} wide.");
        }

        xml.WriteDeclaration();
        xml.WriteRaw("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\""u8);
        xml.WriteInt32(table.Id);
        xml.WriteRaw("\" name=\""u8);
        xml.WriteAttributeText(table.Name);
        xml.WriteRaw("\" displayName=\""u8);
        xml.WriteAttributeText(table.DisplayName ?? table.Name);
        xml.WriteRaw("\" ref=\""u8);
        xml.WriteFormattable(table.Range);
        if (!table.HasHeaderRow)
            xml.WriteRaw("\" headerRowCount=\"0"u8);
        if (table.HasTotalsRow)
            xml.WriteRaw("\" totalsRowCount=\"1"u8);
        xml.WriteRaw("\">"u8);

        // The filter range excludes the totals row: Excel refuses to filter over it.
        int lastFilterRow = table.HasTotalsRow ? table.Range.End.Row - 1 : table.Range.End.Row;
        xml.WriteRaw("<autoFilter ref=\""u8);
        xml.WriteFormattable(new RangeRef(table.Range.Start, new CellRef(lastFilterRow, table.Range.End.Column)));
        xml.WriteRaw("\"/><tableColumns count=\""u8);
        xml.WriteInt32(table.Columns.Count);
        xml.WriteRaw("\">"u8);

        for (int i = 0; i < table.Columns.Count; i++)
        {
            xml.WriteRaw("<tableColumn id=\""u8);
            xml.WriteInt32(i + 1);
            xml.WriteRaw("\" name=\""u8);
            xml.WriteAttributeText(table.Columns[i]);
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</tableColumns>"u8);
        if (table.StyleName is { } style)
        {
            xml.WriteRaw("<tableStyleInfo name=\""u8);
            xml.WriteAttributeText(style);
            xml.WriteRaw("\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\""u8);
            xml.WriteRaw(table.ShowRowStripes ? "1"u8 : "0"u8);
            xml.WriteRaw("\" showColumnStripes=\""u8);
            xml.WriteRaw(table.ShowColumnStripes ? "1"u8 : "0"u8);
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</table>"u8);
    }
}
