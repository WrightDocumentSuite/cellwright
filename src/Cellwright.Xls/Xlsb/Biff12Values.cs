using System.Buffers.Binary;
using System.Text;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// The field types BIFF12 record bodies are built from.
/// </summary>
/// <remarks>
/// Strings are the one place the binary formats really parted ways. BIFF8 wrote a two-byte
/// count and a flag byte saying whether the characters were one byte or two; BIFF12 dropped the
/// choice and writes a four-byte count of characters that are always UTF-16. Numbers did not
/// change at all, so <see cref="BiffValues.DecodeRk"/> serves both.
/// </remarks>
internal static class Biff12Values
{
    /// <summary>The count a nullable string uses to say it holds nothing.</summary>
    private const uint NoString = 0xFFFFFFFF;

    /// <summary>
    /// Reads an <c>XLWideString</c> (MS-XLSB 2.5.169): four bytes of character count, then the
    /// characters. Advances <paramref name="offset"/> past the whole thing.
    /// </summary>
    public static string ReadString(
        ReadOnlySpan<byte> body,
        ref int offset,
        DocumentLoadBudgetState? loadBudget = null) =>
        ReadNullableString(body, ref offset, loadBudget) ?? "";

    /// <summary>
    /// Reads an <c>XLNullableWideString</c> (MS-XLSB 2.5.167), whose count may say that there
    /// is no string at all — which is how an absent relationship id is written, among others.
    /// </summary>
    public static string? ReadNullableString(
        ReadOnlySpan<byte> body,
        ref int offset,
        DocumentLoadBudgetState? loadBudget = null)
    {
        if (offset + 4 > body.Length)
        {
            offset = body.Length;
            return null;
        }

        uint characters = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (characters == NoString)
            return null;

        // A count wider than the record is a corrupt part rather than a huge string, so it is
        // clamped to what is actually there instead of being trusted into an allocation.
        long bytes = Math.Min((long)characters * 2, body.Length - offset);
        if (bytes <= 0)
            return "";

        loadBudget?.EnsureCellTextLength(bytes / 2);
        string text = Encoding.Unicode.GetString(body.Slice(offset, (int)bytes));
        offset += (int)bytes;
        return text;
    }

    /// <summary>
    /// Reads a string only if its stated length fits the record exactly as written, refusing
    /// rather than clamping. Callers that are testing a guess about where a field starts need
    /// the refusal: a clamped read succeeds at any offset and tells them nothing.
    /// </summary>
    public static bool TryReadString(ReadOnlySpan<byte> body, ref int offset, out string? text)
    {
        text = null;
        if (offset + 4 > body.Length)
            return false;

        uint characters = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (characters == NoString)
            return true;

        long bytes = (long)characters * 2;
        if (bytes > body.Length - offset)
            return false;

        text = Encoding.Unicode.GetString(body.Slice(offset, (int)bytes));
        offset += (int)bytes;
        return true;
    }

    /// <summary>
    /// Reads the <c>Cell</c> structure every cell record opens with (MS-XLSB 2.5.10): the
    /// column, and the index of the format that styles it. The row is not in it — that came
    /// from the <c>BrtRowHdr</c> the cells of a row follow.
    /// </summary>
    public static bool TryReadCell(ReadOnlySpan<byte> body, out int column, out int styleIndex)
    {
        column = 0;
        styleIndex = 0;
        if (body.Length < 8)
            return false;

        column = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        styleIndex = (int)(BinaryPrimitives.ReadUInt32LittleEndian(body[4..]) & 0x00FFFFFF);
        return true;
    }

    /// <summary>Maps the one-byte error code of a cell onto the model's error value.</summary>
    public static ExcelError ReadError(byte code) => BiffValues.DecodeError(code);
}
