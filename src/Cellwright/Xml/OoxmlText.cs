using System.Globalization;

namespace Cellwright.Xml;

/// <summary>
/// Decoder for the OOXML <c>_xHHHH_</c> escape convention used to smuggle
/// XML-invalid characters through part content.
/// </summary>
internal static class OoxmlText
{
    /// <summary>Decodes <c>_xHHHH_</c> sequences; returns the original instance when none are present.</summary>
    public static string Decode(string text)
    {
        int start = text.IndexOf("_x", StringComparison.Ordinal);
        if (start < 0 || !ContainsEscape(text, start))
            return text;

        return DecodeCore(text, start);
    }

    private static bool ContainsEscape(string text, int firstCandidate)
    {
        for (int i = firstCandidate; i >= 0 && i + 6 < text.Length; i = text.IndexOf("_x", i + 1, StringComparison.Ordinal))
        {
            if (IsEscapeAt(text, i))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Whether an escape starts here. The <c>x</c> is part of the sequence and has to be
    /// checked: without it a name like <c>A_B1234_C</c> reads as an escape and loses six
    /// characters to a character nobody wrote.
    /// </summary>
    private static bool IsEscapeAt(string text, int index) =>
        index + 6 < text.Length &&
        text[index + 1] == 'x' &&
        char.IsAsciiHexDigit(text[index + 2]) && char.IsAsciiHexDigit(text[index + 3]) &&
        char.IsAsciiHexDigit(text[index + 4]) && char.IsAsciiHexDigit(text[index + 5]) &&
        text[index + 6] == '_';

    private static string DecodeCore(string text, int firstCandidate)
    {
        var builder = new System.Text.StringBuilder(text.Length);
        builder.Append(text, 0, firstCandidate);
        int position = firstCandidate;
        while (position < text.Length)
        {
            if (text[position] == '_' && IsEscapeAt(text, position))
            {
                int code = int.Parse(text.AsSpan(position + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                builder.Append((char)code);
                position += 7;
            }
            else
            {
                builder.Append(text[position]);
                position++;
            }
        }

        return builder.ToString();
    }
}
