using System.Diagnostics;

namespace Cellwright;

/// <summary>
/// A 1-based reference to a single cell, e.g. <c>B7</c>. Parsing accepts and ignores
/// <c>$</c> anchors, so <c>$B$7</c> resolves to the same reference.
/// </summary>
[DebuggerDisplay("{ToString(),nq}")]
public readonly record struct CellRef(int Row, int Column) :
    ISpanParsable<CellRef>, ISpanFormattable, IUtf8SpanFormattable
{
    /// <summary>Longest textual form: three column letters plus seven row digits.</summary>
    public const int MaxTextLength = 10;

    /// <summary>Returns <see langword="true"/> when both coordinates lie inside the worksheet grid.</summary>
    public bool IsValid => SheetLimits.IsValidRow(Row) && SheetLimits.IsValidColumn(Column);

    /// <summary>Parses an A1-style reference such as <c>C12</c> or <c>$C$12</c>.</summary>
    /// <exception cref="FormatException">The text is not a valid cell reference.</exception>
    public static CellRef Parse(ReadOnlySpan<char> s, IFormatProvider? provider = null)
        => TryParse(s, provider, out CellRef result) ? result : throw new FormatException($"'{s}' is not a valid A1 cell reference.");

    /// <inheritdoc cref="Parse(ReadOnlySpan{char}, IFormatProvider?)"/>
    public static CellRef Parse(string s, IFormatProvider? provider = null)
    {
        ArgumentNullException.ThrowIfNull(s);
        return Parse(s.AsSpan(), provider);
    }

    /// <summary>Attempts to parse an A1-style reference without throwing.</summary>
    public static bool TryParse(ReadOnlySpan<char> s, IFormatProvider? provider, out CellRef result)
    {
        result = default;
        int position = 0;
        if (!ColumnName.TryParse(s, ref position, out int column))
            return false;

        if (position < s.Length && s[position] == '$')
            position++;

        ReadOnlySpan<char> digits = s[position..];
        if (digits.IsEmpty || !int.TryParse(digits, System.Globalization.NumberStyles.None, provider, out int row))
            return false;
        if (!SheetLimits.IsValidRow(row))
            return false;

        result = new CellRef(row, column);
        return true;
    }

    /// <inheritdoc cref="TryParse(ReadOnlySpan{char}, IFormatProvider?, out CellRef)"/>
    public static bool TryParse(string? s, IFormatProvider? provider, out CellRef result)
    {
        result = default;
        return s is not null && TryParse(s.AsSpan(), provider, out result);
    }

    /// <summary>Formats the reference as A1 text, e.g. <c>AB12</c>.</summary>
    public override string ToString()
    {
        Span<char> buffer = stackalloc char[MaxTextLength];
        TryFormat(buffer, out int written, default, null);
        return new string(buffer[..written]);
    }

    string IFormattable.ToString(string? format, IFormatProvider? formatProvider) => ToString();

    /// <inheritdoc/>
    public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
    {
        charsWritten = 0;
        if (!ColumnName.TryFormat(Column, destination, out int letters))
            return false;
        if (!Row.TryFormat(destination[letters..], out int digits, default, System.Globalization.CultureInfo.InvariantCulture))
            return false;
        charsWritten = letters + digits;
        return true;
    }

    /// <inheritdoc/>
    public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
    {
        bytesWritten = 0;
        if (!ColumnName.TryFormatUtf8(Column, utf8Destination, out int letters))
            return false;
        if (!Row.TryFormat(utf8Destination[letters..], out int digits, default, System.Globalization.CultureInfo.InvariantCulture))
            return false;
        bytesWritten = letters + digits;
        return true;
    }
}
