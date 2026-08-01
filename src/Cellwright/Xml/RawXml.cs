using System.Xml;

namespace Cellwright.Xml;

/// <summary>
/// Capture and replay of markup the model does not interpret. Attributes are kept as decoded
/// name/value pairs (re-escaped on the way out); elements are kept as the verbatim XML the
/// reader produced, so anything from a theme override to a pivot cache survives a round trip.
/// </summary>
internal static class RawXml
{
    /// <summary>Attributes of the current element, excluding the given names. Returns <see langword="null"/> when nothing is left.</summary>
    public static List<KeyValuePair<string, string>>? ReadAttributes(XmlReader xml, params ReadOnlySpan<string> skip)
    {
        if (!xml.HasAttributes)
            return null;

        List<KeyValuePair<string, string>>? attributes = null;
        if (xml.MoveToFirstAttribute())
        {
            do
            {
                if (skip.Contains(xml.Name))
                    continue;
                (attributes ??= []).Add(new KeyValuePair<string, string>(xml.Name, xml.Value));
            }
            while (xml.MoveToNextAttribute());

            xml.MoveToElement();
        }

        return attributes;
    }

    /// <summary>Writes captured attributes, each preceded by a space, ready to sit inside a start tag.</summary>
    public static void WriteAttributes(Utf8XmlWriter writer, List<KeyValuePair<string, string>>? attributes)
    {
        if (attributes is null)
            return;

        foreach ((string name, string value) in attributes)
        {
            writer.WriteRaw(" "u8);
            writer.WriteRawXml(name);
            writer.WriteRaw("=\""u8);
            writer.WriteAttributeText(value);
            writer.WriteRaw("\""u8);
        }
    }

    /// <summary>Writes a captured element fragment verbatim.</summary>
    public static void Write(Utf8XmlWriter writer, string? fragment)
    {
        if (!string.IsNullOrEmpty(fragment))
            writer.WriteRawXml(fragment);
    }

    /// <summary>Appends the current element's markup to a buffer, leaving the reader on the following node.</summary>
    public static async ValueTask<System.Text.StringBuilder> AppendOuterXmlAsync(
        System.Text.StringBuilder? buffer, XmlReader xml)
    {
        buffer ??= new System.Text.StringBuilder();
        buffer.Append(await xml.ReadOuterXmlAsync().ConfigureAwait(false));
        return buffer;
    }
}
