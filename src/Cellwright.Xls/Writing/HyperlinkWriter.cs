using System.Text;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the HLINK record. The old format describes a link's target with an OLE moniker rather
/// than a plain string, so a URL and a file path are encoded differently.
/// </summary>
internal static class HyperlinkWriter
{
    private static ReadOnlySpan<byte> HyperlinkClassId =>
        [0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B];

    private static ReadOnlySpan<byte> UrlMoniker =>
        [0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B];

    /// <summary>Writes one hyperlink.</summary>
    public static void Write(BiffWriter writer, Hyperlink hyperlink)
    {
        RangeRef range = hyperlink.Range;
        int start = writer.BeginRecord(BiffRecordType.HLink);
        writer.AppendUInt16(range.Start.Row - 1);
        writer.AppendUInt16(range.End.Row - 1);
        writer.AppendUInt16(range.Start.Column - 1);
        writer.AppendUInt16(range.End.Column - 1);
        writer.Append(HyperlinkClassId);
        writer.AppendUInt32(2);                   // Stream version.

        bool hasTarget = !string.IsNullOrEmpty(hyperlink.Target);
        bool hasLocation = !string.IsNullOrEmpty(hyperlink.Location);
        bool hasDisplay = !string.IsNullOrEmpty(hyperlink.Display);

        // MS-OSHARED 2.3.7.1. Bit 4 is what says the display text is present; bit 2 only says
        // the author chose it. The tooltip is not in this structure at all.
        uint flags = 0;
        if (hasTarget)
            flags |= 0x0001 | 0x0002;             // A target, stated absolutely.
        if (hasLocation)
            flags |= 0x0008;
        if (hasDisplay)
            flags |= 0x0010 | 0x0004;
        writer.AppendUInt32(flags);

        if (hasDisplay)
            AppendWideString(writer, hyperlink.Display!);

        if (hasTarget)
        {
            // Every target is written as a URL moniker: it is the one form that carries an
            // arbitrary string, and Excel resolves a relative path in it just the same.
            writer.Append(UrlMoniker);
            byte[] bytes = Encoding.Unicode.GetBytes(hyperlink.Target + "\0");
            writer.AppendUInt32((uint)bytes.Length);
            writer.Append(bytes);
        }

        if (hasLocation)
            AppendWideString(writer, hyperlink.Location!);

        writer.EndRecord(start);

        if (!string.IsNullOrEmpty(hyperlink.Tooltip))
            WriteTooltip(writer, range, hyperlink.Tooltip);
    }

    /// <summary>
    /// Writes the HLinkTooltip that follows the link (MS-XLS 2.4.141). It repeats the record
    /// type inside its own body and names the range again, which is how Excel pairs the two.
    /// </summary>
    private static void WriteTooltip(BiffWriter writer, RangeRef range, string tooltip)
    {
        int start = writer.BeginRecord(BiffRecordType.HLinkTooltip);
        writer.AppendUInt16(BiffRecordType.HLinkTooltip);
        writer.AppendUInt16(range.Start.Row - 1);
        writer.AppendUInt16(range.End.Row - 1);
        writer.AppendUInt16(range.Start.Column - 1);
        writer.AppendUInt16(range.End.Column - 1);

        // The string is null-terminated and, unlike a HyperlinkString, carries no count.
        foreach (char c in tooltip)
            writer.AppendUInt16(c);
        writer.AppendUInt16(0);
        writer.EndRecord(start);
    }

    /// <summary>Writes a counted UTF-16 string whose count includes the terminating NUL.</summary>
    private static void AppendWideString(BiffWriter writer, string text)
    {
        writer.AppendUInt32((uint)(text.Length + 1));
        foreach (char c in text)
            writer.AppendUInt16(c);
        writer.AppendUInt16(0);
    }
}
