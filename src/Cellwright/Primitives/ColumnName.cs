namespace Cellwright;

/// <summary>
/// Conversions between 1-based column numbers and their letter names (<c>A</c>..<c>XFD</c>).
/// </summary>
public static class ColumnName
{
    /// <summary>Converts a column number to letters, e.g. 28 to <c>AB</c>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column lies outside the grid.</exception>
    public static string Format(int column)
    {
        Span<char> buffer = stackalloc char[3];
        if (!TryFormat(column, buffer, out int written))
            throw new ArgumentOutOfRangeException(nameof(column), column, "Column number must be between 1 and 16384.");
        return new string(buffer[..written]);
    }

    /// <summary>Writes the letter name of <paramref name="column"/> into <paramref name="destination"/>.</summary>
    public static bool TryFormat(int column, Span<char> destination, out int charsWritten)
    {
        charsWritten = 0;
        if (!SheetLimits.IsValidColumn(column))
            return false;

        Span<char> scratch = stackalloc char[3];
        int count = 0;
        while (column > 0)
        {
            column--;
            scratch[count++] = (char)('A' + column % 26);
            column /= 26;
        }

        if (destination.Length < count)
            return false;
        for (int i = 0; i < count; i++)
            destination[i] = scratch[count - 1 - i];
        charsWritten = count;
        return true;
    }

    /// <summary>UTF-8 variant of <see cref="TryFormat(int, Span{char}, out int)"/>.</summary>
    public static bool TryFormatUtf8(int column, Span<byte> destination, out int bytesWritten)
    {
        Span<char> chars = stackalloc char[3];
        if (!TryFormat(column, chars, out int count) || destination.Length < count)
        {
            bytesWritten = 0;
            return false;
        }

        for (int i = 0; i < count; i++)
            destination[i] = (byte)chars[i];
        bytesWritten = count;
        return true;
    }

    /// <summary>Parses a full letter name, e.g. <c>AB</c> to 28.</summary>
    /// <exception cref="FormatException">The text is not a valid column name.</exception>
    public static int Parse(ReadOnlySpan<char> name)
    {
        int position = 0;
        if (!TryParse(name, ref position, out int column) || position != name.Length)
            throw new FormatException($"'{name}' is not a valid column name.");
        return column;
    }

    /// <summary>
    /// Reads a column name starting at <paramref name="position"/>, ignoring a leading <c>$</c> anchor.
    /// Advances <paramref name="position"/> past the letters. Accepts both letter cases.
    /// </summary>
    public static bool TryParse(ReadOnlySpan<char> text, ref int position, out int column)
    {
        column = 0;
        int i = position;
        if (i < text.Length && text[i] == '$')
            i++;

        int value = 0;
        int letters = 0;
        while (i < text.Length && letters < 3)
        {
            int digit = (text[i] | 0x20) - 'a'; // Case-insensitive: fold to lowercase.
            if ((uint)digit >= 26)
                break;
            value = value * 26 + digit + 1;
            letters++;
            i++;
        }

        if (letters == 0 || value > SheetLimits.MaxColumn)
            return false;

        column = value;
        position = i;
        return true;
    }
}
