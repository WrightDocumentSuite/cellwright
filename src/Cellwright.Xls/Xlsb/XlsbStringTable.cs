using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the shared string table: the text of every string cell in the workbook, stored once
/// and pointed at by index.
/// </summary>
internal static class XlsbStringTable
{
    public static void Read(byte[] part, XlsGlobals globals)
    {
        var strings = new List<string>();
        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            if (reader.RecordType == Biff12RecordType.SstItem)
                strings.Add(ReadRichString(reader.Record));
        }

        globals.SharedStrings = [.. strings];
    }

    /// <summary>
    /// Reads a <c>RichStr</c> (MS-XLSB 2.5.122): a byte of flags, the text, and then the runs
    /// of formatting over it. The runs are skipped — the model holds one style per cell, so a
    /// string formatted in pieces arrives as plain text, as it does from xlsx and xls.
    /// </summary>
    public static string ReadRichString(ReadOnlySpan<byte> body)
    {
        if (body.Length < 1)
            return "";

        int offset = 1;
        return Biff12Values.ReadString(body, ref offset);
    }
}
