using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Styles;

/// <summary>
/// The children of a font element, read and written in one place.
/// </summary>
/// <remarks>
/// The format spells a font twice: as <c>font</c> in the styles part, and as <c>rPr</c> on a run
/// of formatted text inside a string. The two carry the same children under the same names, apart
/// from the family, which is <c>name</c> in one and <c>rFont</c> in the other. Keeping one
/// implementation is what stops a property from being understood in a cell's font and quietly
/// dropped from a run's — the two would otherwise drift apart a property at a time.
/// <para>
/// Order is not a concern: the schema declares both as an unbounded choice (CT_Font and
/// CT_RPrElt, §18.8.22 and §18.4.7), so children may appear in any order.
/// </para>
/// </remarks>
internal static class FontMarkup
{
    /// <summary>Writes the children of a font element, without the element itself.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="font">The font to write.</param>
    /// <param name="runProperties">Whether this is an <c>rPr</c> rather than a <c>font</c>.</param>
    public static void WriteChildren(Utf8XmlWriter xml, FontStyle font, bool runProperties)
    {
        if (font.Bold)
            xml.WriteRaw("<b/>"u8);
        if (font.Italic)
            xml.WriteRaw("<i/>"u8);
        if (font.Strikethrough)
            xml.WriteRaw("<strike/>"u8);

        StylesPartWriter.WriteUnderline(xml, font.Underline);

        if (font.VerticalAlign != FontVerticalAlign.Baseline)
        {
            xml.WriteRaw("<vertAlign val=\""u8);
            xml.WriteRaw(font.VerticalAlign == FontVerticalAlign.Superscript ? "superscript"u8 : "subscript"u8);
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("<sz val=\""u8);
        xml.WriteDouble(font.Size);
        xml.WriteRaw("\"/>"u8);

        StylesPartWriter.WriteColorElement(xml, "<color rgb=\""u8, font.Color);

        xml.WriteRaw(runProperties ? "<rFont val=\""u8 : "<name val=\""u8);
        xml.WriteAttributeText(font.Name);
        xml.WriteRaw("\"/>"u8);

        StylesPartWriter.WriteIntElement(xml, "<family val=\""u8, font.Family, font.Family > 0);
        StylesPartWriter.WriteIntElement(xml, "<charset val=\""u8, font.Charset, font.Charset >= 0);

        if (font.Scheme != FontScheme.None)
        {
            xml.WriteRaw("<scheme val=\""u8);
            xml.WriteRaw(font.Scheme == FontScheme.Major ? "major"u8 : "minor"u8);
            xml.WriteRaw("\"/>"u8);
        }
    }

    /// <summary>
    /// Reads a font element. The reader starts on the element and ends on its end tag, or on the
    /// element itself when it is empty.
    /// </summary>
    /// <param name="xml">The reader, positioned on the element.</param>
    /// <param name="container">Local name of the element being read: <c>font</c> or <c>rPr</c>.</param>
    public static FontStyle Parse(XmlReader xml, string container)
    {
        FontStyle font = FontStyle.Default;
        if (xml.IsEmptyElement)
            return font;

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == container)
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            font = xml.LocalName switch
            {
                "b" => font with { Bold = StylesPartParser.IsFlagSet(xml) },
                "i" => font with { Italic = StylesPartParser.IsFlagSet(xml) },
                "strike" => font with { Strikethrough = StylesPartParser.IsFlagSet(xml) },
                "u" => font with { Underline = StylesPartParser.ParseUnderline(xml.GetAttribute("val")) },
                "vertAlign" => font with { VerticalAlign = StylesPartParser.ParseVerticalAlign(xml.GetAttribute("val")) },
                "sz" when double.TryParse(
                    xml.GetAttribute("val"), System.Globalization.CultureInfo.InvariantCulture, out double size)
                    => font with { Size = size },
                "color" => font with { Color = StylesPartParser.ParseColor(xml) },

                // The one place the two spellings differ.
                "name" or "rFont" when xml.GetAttribute("val") is { } name => font with { Name = name },
                "family" => font with { Family = StylesPartParser.ParseIntAttribute(xml, "val") },
                "charset" => font with { Charset = StylesPartParser.ParseIntAttribute(xml, "val") },
                "scheme" => font with { Scheme = StylesPartParser.ParseScheme(xml.GetAttribute("val")) },
                _ => font,
            };
        }

        return font;
    }
}
