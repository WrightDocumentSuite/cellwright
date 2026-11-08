using System.Buffers.Binary;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the shared string table: the text of every string cell in the workbook, stored once
/// and pointed at by index.
/// </summary>
internal static class XlsbStringTable
{
    public static void Read(
        byte[] part,
        XlsGlobals globals,
        DocumentLoadBudgetState? loadBudget = null,
        CancellationToken cancellationToken = default)
    {
        var strings = new List<string>();
        var richStrings = new List<RichText?>();
        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.RecordType == Biff12RecordType.SstItem)
            {
                loadBudget?.AddSharedString();
                string text = ReadRichString(reader.Record, globals, out RichText? rich, loadBudget);
                loadBudget?.AddCellText(text);
                strings.Add(text);
                richStrings.Add(rich);
            }
        }

        globals.SharedStrings = [.. strings];
        globals.SharedRichStrings = [.. richStrings];
    }

    /// <summary>
    /// Reads a <c>RichStr</c> (MS-XLSB 2.5.122): a byte of flags, the text, and any formatting runs.
    /// </summary>
    public static string ReadRichString(
        ReadOnlySpan<byte> body,
        DocumentLoadBudgetState? loadBudget = null)
    {
        if (body.Length < 1)
            return "";

        int offset = 1;
        return Biff12Values.ReadString(body, ref offset, loadBudget);
    }

    /// <summary>
    /// Reads a <c>RichStr</c> (MS-XLSB 2.5.122) along with its character-formatting runs (<c>StrRun</c>).
    /// </summary>
    public static string ReadRichString(
        ReadOnlySpan<byte> body,
        XlsGlobals globals,
        out RichText? richText,
        DocumentLoadBudgetState? loadBudget = null)
    {
        richText = null;
        if (body.Length < 1)
            return "";

        byte flags = body[0];
        int offset = 1;
        string text = Biff12Values.ReadString(body, ref offset, loadBudget);
        if ((flags & 0x01) != 0 && offset + 4 <= body.Length)
        {
            uint runCount = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
            offset += 4;
            int available = (int)Math.Min(runCount, (uint)((body.Length - offset) / 4));
            if (available > 0)
            {
                var runs = new (int CharPos, int FontIndex)[available];
                for (int i = 0; i < available; i++)
                {
                    runs[i] = (
                        BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]),
                        BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]));
                    offset += 4;
                }

                richText = globals.BuildRichText(text, runs);
            }
        }

        return text;
    }
}
