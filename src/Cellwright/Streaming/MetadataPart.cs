using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// The cell metadata part, as far as dynamic arrays need it.
/// </summary>
/// <remarks>
/// The part is a level of indirection: a cell's <c>cm</c> attribute indexes a list of records,
/// each record names a metadata type, and one of those types — <c>XLDAPR</c>, carrying
/// <c>xda:dynamicArrayProperties</c> — is what marks a formula as a dynamic array rather than a
/// legacy CSE one. Both spell their formula the same way (<c>t="array"</c> with a <c>ref</c>),
/// so without this part the two are indistinguishable, and a reader that ignored it would turn
/// every spill into a CSE formula on round-trip.
/// </remarks>
internal static class MetadataPart
{
    /// <summary>Namespace of the dynamic array extension (MS-XLSX 2.3.7.4).</summary>
    private const string DynamicArrayNamespace =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

    /// <summary>Extension identifier Excel writes for the dynamic array properties.</summary>
    private const string DynamicArrayExtension = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";

    /// <summary>
    /// Reads which <c>cm</c> indexes mark a dynamic array: the 1-based positions in the
    /// <c>cellMetadata</c> list whose record points at the <c>XLDAPR</c> type.
    /// </summary>
    /// <param name="stream">The part's content.</param>
    public static HashSet<int> ReadDynamicArrayIndexes(Stream stream)
    {
        var dynamicIndexes = new HashSet<int>();
        var typeNames = new List<string>();
        int position = 0;

        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        bool inCellMetadata = false;
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.Element)
            {
                switch (xml.LocalName)
                {
                    case "metadataType" when xml.GetAttribute("name") is { } name:
                        typeNames.Add(name);
                        break;

                    case "cellMetadata":
                        inCellMetadata = true;
                        break;

                    case "rc" when inCellMetadata:
                    {
                        position++;
                        if (int.TryParse(xml.GetAttribute("t"), NumberStyles.None, CultureInfo.InvariantCulture, out int type) &&
                            type >= 1 && type <= typeNames.Count &&
                            typeNames[type - 1] == "XLDAPR")
                        {
                            dynamicIndexes.Add(position);
                        }

                        break;
                    }

                    default:
                        break;
                }
            }
            else if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "cellMetadata")
            {
                inCellMetadata = false;
            }
        }

        return dynamicIndexes;
    }

    /// <summary>
    /// Writes the smallest metadata part every dynamic array anchor can share: one type, one
    /// record, so every anchor's <c>cm</c> is 1.
    /// </summary>
    /// <param name="xml">Where to write.</param>
    /// <remarks>
    /// The flags on the type are the ones Excel writes; they say what happens to the mark when
    /// the cell is copied, pasted or shifted, and a reader that honours them expects exactly
    /// this set for a dynamic array.
    /// </remarks>
    public static void Write(Utf8XmlWriter xml)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""u8);
        xml.WriteRaw(" xmlns:xda=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(DynamicArrayNamespace));
        xml.WriteRaw("\">"u8);

        xml.WriteRaw("<metadataTypes count=\"1\">"u8);
        xml.WriteRaw("<metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\""u8);
        xml.WriteRaw(" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\""u8);
        xml.WriteRaw(" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>"u8);
        xml.WriteRaw("</metadataTypes>"u8);

        xml.WriteRaw("<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(DynamicArrayExtension));
        xml.WriteRaw("\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>"u8);

        xml.WriteRaw("<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>"u8);
        xml.WriteRaw("</metadata>"u8);
    }
}
