using System.Globalization;
using System.Text;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// The three parts under <c>docProps</c>: what the document says about itself, what the program
/// that wrote it recorded, and whatever the author added on top. None of them affects a single
/// cell, which is why they are easy to lose — and why losing them is noticed only much later.
/// </summary>
internal static class DocumentPropertiesPart
{
    private const string CoreNamespaces =
        " xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"" +
        " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"" +
        " xmlns:dcterms=\"http://purl.org/dc/terms/\"" +
        " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"" +
        " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"";

    private const string VariantNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>The identifier every custom property carries; it is the same one for all of them.</summary>
    private const string CustomFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    /// <summary>Reads <c>core.xml</c> into the model, keeping the fields it does not name.</summary>
    public static void ReadCore(Stream stream, WorkbookProperties properties)
    {
        StringBuilder? unmodelled = null;
        using XmlReader xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);

        // Reading an element's content steps past its end tag, so the loop must not advance
        // again or the element after it goes unseen.
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;

            advance = true;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "coreProperties":
                    continue;
                case "title": properties.Title = Text(xml, ref advance); break;
                case "subject": properties.Subject = Text(xml, ref advance); break;
                case "creator": properties.Author = Text(xml, ref advance); break;
                case "keywords": properties.Keywords = Text(xml, ref advance); break;
                case "description": properties.Description = Text(xml, ref advance); break;
                case "category": properties.Category = Text(xml, ref advance); break;
                case "contentStatus": properties.ContentStatus = Text(xml, ref advance); break;
                case "lastModifiedBy": properties.LastModifiedBy = Text(xml, ref advance); break;
                case "created": properties.Created = ParseTimestamp(Text(xml, ref advance)); break;
                case "modified": properties.SetLoadedModified(ParseTimestamp(Text(xml, ref advance))); break;
                default:
                    unmodelled = Capture(unmodelled, xml, ref advance);
                    break;
            }
        }

        properties.CoreMarkup = unmodelled?.ToString();
    }

    /// <summary>Reads <c>app.xml</c> into the model, keeping the fields it does not name.</summary>
    public static void ReadExtended(Stream stream, WorkbookProperties properties)
    {
        StringBuilder? unmodelled = null;
        using XmlReader xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;

            advance = true;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "Properties":
                    continue;
                case "Application": properties.Application = Text(xml, ref advance); break;
                case "Company": properties.Company = Text(xml, ref advance); break;
                case "Manager": properties.Manager = Text(xml, ref advance); break;

                // Both of these list the sheets by name, and the model is free to have renamed,
                // added or removed one since. They are rebuilt on save rather than replayed.
                case "HeadingPairs" or "TitlesOfParts":
                    xml.Skip();
                    advance = false;
                    break;
                default:
                    unmodelled = Capture(unmodelled, xml, ref advance);
                    break;
            }
        }

        properties.ExtendedMarkup = unmodelled?.ToString();
    }

    /// <summary>
    /// The text of the current element. An empty one is left where it is, so that the loop
    /// steps past it; a full one is consumed along with its end tag, so the loop must not.
    /// </summary>
    private static string Text(XmlReader xml, ref bool advance)
    {
        if (xml.IsEmptyElement)
            return string.Empty;

        advance = false;
        return xml.ReadElementContentAsString();
    }

    private static StringBuilder Capture(StringBuilder? buffer, XmlReader xml, ref bool advance)
    {
        advance = false;
        return (buffer ??= new StringBuilder()).Append(xml.ReadOuterXml());
    }

    /// <summary>Reads <c>custom.xml</c>, mapping each variant onto the value kind it means.</summary>
    public static void ReadCustom(Stream stream, WorkbookProperties properties)
    {
        using XmlReader xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "property")
                continue;

            if (xml.GetAttribute("name") is not { Length: > 0 } name || xml.IsEmptyElement)
                continue;

            if (ReadVariant(xml) is { IsBlank: false } value)
                properties.Custom[name] = value;
        }
    }

    private static CellValue ReadVariant(XmlReader xml)
    {
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "property")
                break;
            if (xml.NodeType != XmlNodeType.Element || xml.NamespaceURI != VariantNamespace)
                continue;

            string kind = xml.LocalName;
            string text = xml.ReadElementContentAsString();
            return kind switch
            {
                "lpwstr" or "lpstr" or "bstr" => CellValue.Text(text),
                "i1" or "i2" or "i4" or "i8" or "int" or "ui1" or "ui2" or "ui4" or "ui8" or "uint" or "r4" or "r8" or "decimal" =>
                    double.TryParse(text, CultureInfo.InvariantCulture, out double number)
                        ? CellValue.Number(number)
                        : CellValue.Text(text),
                "bool" => CellValue.Boolean(text is "true" or "1"),
                "filetime" or "date" => ParseTimestamp(text) is { } date ? CellValue.DateTime(date) : CellValue.Text(text),
                _ => CellValue.Text(text),
            };
        }

        return CellValue.Blank;
    }

    /// <summary>Writes <c>core.xml</c>, stamping the save time unless the caller chose one.</summary>
    public static void WriteCore(Utf8XmlWriter xml, WorkbookProperties properties)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<cp:coreProperties"u8);
        xml.WriteRawXml(CoreNamespaces);
        xml.WriteRaw(">"u8);

        WriteElement(xml, "dc:title", properties.Title);
        WriteElement(xml, "dc:subject", properties.Subject);
        WriteElement(xml, "dc:creator", properties.Author);
        WriteElement(xml, "cp:keywords", properties.Keywords);
        WriteElement(xml, "dc:description", properties.Description);
        WriteElement(xml, "cp:lastModifiedBy", properties.LastModifiedBy);
        WriteTimestamp(xml, "dcterms:created", properties.Created ?? DateTime.UtcNow);
        WriteTimestamp(xml, "dcterms:modified", properties.ModifiedOnSave());
        WriteElement(xml, "cp:category", properties.Category);
        WriteElement(xml, "cp:contentStatus", properties.ContentStatus);
        RawXml.Write(xml, properties.CoreMarkup);

        xml.WriteRaw("</cp:coreProperties>"u8);
    }

    /// <summary>Writes <c>app.xml</c>, rebuilding the sheet list from the workbook as it stands.</summary>
    public static void WriteExtended(Utf8XmlWriter xml, WorkbookProperties properties, IReadOnlyList<string> sheetNames)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\""u8);
        xml.WriteRaw(" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"u8);

        WriteElement(xml, "Application", properties.Application ?? "Cellwright");
        WriteElement(xml, "Manager", properties.Manager);
        WriteElement(xml, "Company", properties.Company);
        WriteSheetNames(xml, sheetNames);
        RawXml.Write(xml, properties.ExtendedMarkup);

        xml.WriteRaw("</Properties>"u8);
    }

    /// <summary>Writes <c>custom.xml</c>; the caller only calls this when there is something in it.</summary>
    public static void WriteCustom(Utf8XmlWriter xml, WorkbookProperties properties)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\""u8);
        xml.WriteRaw(" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"u8);

        // Identifiers start at two: one is reserved for the dictionary the format never uses.
        int id = 2;
        foreach ((string name, CellValue value) in properties.Custom)
        {
            xml.WriteRaw("<property fmtid=\""u8);
            xml.WriteRawXml(CustomFormatId);
            xml.WriteRaw("\" pid=\""u8);
            xml.WriteInt32(id++);
            xml.WriteRaw("\" name=\""u8);
            xml.WriteAttributeText(name);
            xml.WriteRaw("\">"u8);
            WriteVariant(xml, value);
            xml.WriteRaw("</property>"u8);
        }

        xml.WriteRaw("</Properties>"u8);
    }

    private static void WriteVariant(Utf8XmlWriter xml, CellValue value)
    {
        switch (value.Kind)
        {
            case CellValueKind.Number:
                xml.WriteRaw("<vt:r8>"u8);
                xml.WriteDouble(value.GetNumber());
                xml.WriteRaw("</vt:r8>"u8);
                break;
            case CellValueKind.Boolean:
                xml.WriteRaw(value.GetBoolean() ? "<vt:bool>true</vt:bool>"u8 : "<vt:bool>false</vt:bool>"u8);
                break;
            case CellValueKind.DateTime:
                xml.WriteRaw("<vt:filetime>"u8);
                xml.WriteRawXml(Format(value.GetDateTime()));
                xml.WriteRaw("</vt:filetime>"u8);
                break;
            default:
                xml.WriteRaw("<vt:lpwstr>"u8);
                xml.WriteText(value.IsText ? value.GetText() : value.ToString());
                xml.WriteRaw("</vt:lpwstr>"u8);
                break;
        }
    }

    /// <summary>
    /// The two elements that name the sheets: a vector saying how many worksheets there are,
    /// and a vector of their titles. Excel shows them in the document properties pane.
    /// </summary>
    private static void WriteSheetNames(Utf8XmlWriter xml, IReadOnlyList<string> sheetNames)
    {
        if (sheetNames.Count == 0)
            return;

        xml.WriteRaw("<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">"u8);
        xml.WriteRaw("<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"u8);
        xml.WriteRaw("<vt:variant><vt:i4>"u8);
        xml.WriteInt32(sheetNames.Count);
        xml.WriteRaw("</vt:i4></vt:variant></vt:vector></HeadingPairs>"u8);

        xml.WriteRaw("<TitlesOfParts><vt:vector size=\""u8);
        xml.WriteInt32(sheetNames.Count);
        xml.WriteRaw("\" baseType=\"lpstr\">"u8);
        foreach (string name in sheetNames)
        {
            xml.WriteRaw("<vt:lpstr>"u8);
            xml.WriteText(name);
            xml.WriteRaw("</vt:lpstr>"u8);
        }

        xml.WriteRaw("</vt:vector></TitlesOfParts>"u8);
    }

    private static void WriteElement(Utf8XmlWriter xml, string name, string? value)
    {
        if (value is null)
            return;

        xml.WriteRaw("<"u8);
        xml.WriteRawXml(name);
        xml.WriteRaw(">"u8);
        xml.WriteText(value);
        xml.WriteRaw("</"u8);
        xml.WriteRawXml(name);
        xml.WriteRaw(">"u8);
    }

    private static void WriteTimestamp(Utf8XmlWriter xml, string name, DateTime value)
    {
        xml.WriteRaw("<"u8);
        xml.WriteRawXml(name);
        xml.WriteRaw(" xsi:type=\"dcterms:W3CDTF\">"u8);
        xml.WriteRawXml(Format(value));
        xml.WriteRaw("</"u8);
        xml.WriteRawXml(name);
        xml.WriteRaw(">"u8);
    }

    /// <summary>W3C date-time, which is what both the core part and a filetime variant use.</summary>
    private static string Format(DateTime value)
    {
        DateTime utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : value;
        return utc.ToString("yyyy-MM-ddTHH:mm:ss'Z'", CultureInfo.InvariantCulture);
    }

    private static DateTime? ParseTimestamp(string text) =>
        DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime value)
            ? value
            : null;
}
