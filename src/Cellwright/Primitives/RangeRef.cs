using System.Diagnostics;

namespace Cellwright;

/// <summary>
/// A rectangular, inclusive range of cells such as <c>A1:C5</c>. Always normalized:
/// <see cref="Start"/> is the top-left corner and <see cref="End"/> the bottom-right one.
/// </summary>
[DebuggerDisplay("{ToString(),nq}")]
public readonly record struct RangeRef : ISpanParsable<RangeRef>, ISpanFormattable, IUtf8SpanFormattable
{
    /// <summary>Longest textual form: two cell references and a colon.</summary>
    public const int MaxTextLength = CellRef.MaxTextLength * 2 + 1;

    /// <summary>Creates a range from two arbitrary corners; the corners are normalized.</summary>
    public RangeRef(CellRef first, CellRef second)
    {
        Start = new CellRef(Math.Min(first.Row, second.Row), Math.Min(first.Column, second.Column));
        End = new CellRef(Math.Max(first.Row, second.Row), Math.Max(first.Column, second.Column));
    }

    /// <summary>Creates a single-cell range.</summary>
    public RangeRef(CellRef cell) : this(cell, cell)
    {
    }

    /// <summary>Top-left corner.</summary>
    public CellRef Start { get; }

    /// <summary>Bottom-right corner (inclusive).</summary>
    public CellRef End { get; }

    /// <summary>Number of rows in the range.</summary>
    public int RowCount => End.Row - Start.Row + 1;

    /// <summary>Number of columns in the range.</summary>
    public int ColumnCount => End.Column - Start.Column + 1;

    /// <summary>Total number of cells in the range.</summary>
    public long CellCount => (long)RowCount * ColumnCount;

    /// <summary>Returns <see langword="true"/> when the reference covers exactly one cell.</summary>
    public bool IsSingleCell => Start == End;

    /// <summary>Returns <see langword="true"/> when <paramref name="cell"/> lies inside the range.</summary>
    public bool Contains(CellRef cell) =>
        cell.Row >= Start.Row && cell.Row <= End.Row &&
        cell.Column >= Start.Column && cell.Column <= End.Column;

    /// <summary>Returns <see langword="true"/> when the two ranges share at least one cell.</summary>
    public bool Intersects(RangeRef other) =>
        Start.Row <= other.End.Row && End.Row >= other.Start.Row &&
        Start.Column <= other.End.Column && End.Column >= other.Start.Column;

    /// <summary>Parses <c>A1:C5</c> or a single-cell form like <c>B2</c>.</summary>
    /// <exception cref="FormatException">The text is not a valid range reference.</exception>
    public static RangeRef Parse(ReadOnlySpan<char> s, IFormatProvider? provider = null)
        => TryParse(s, provider, out RangeRef result) ? result : throw new FormatException($"'{s}' is not a valid A1 range reference.");

    /// <inheritdoc cref="Parse(ReadOnlySpan{char}, IFormatProvider?)"/>
    public static RangeRef Parse(string s, IFormatProvider? provider = null)
    {
        ArgumentNullException.ThrowIfNull(s);
        return Parse(s.AsSpan(), provider);
    }

    /// <summary>Attempts to parse a range reference without throwing.</summary>
    public static bool TryParse(ReadOnlySpan<char> s, IFormatProvider? provider, out RangeRef result)
    {
        result = default;
        int colon = s.IndexOf(':');
        if (colon < 0)
        {
            if (!CellRef.TryParse(s, provider, out CellRef single))
                return false;
            result = new RangeRef(single);
            return true;
        }

        if (!CellRef.TryParse(s[..colon], provider, out CellRef start) ||
            !CellRef.TryParse(s[(colon + 1)..], provider, out CellRef end))
        {
            return false;
        }

        result = new RangeRef(start, end);
        return true;
    }

    /// <inheritdoc cref="TryParse(ReadOnlySpan{char}, IFormatProvider?, out RangeRef)"/>
    public static bool TryParse(string? s, IFormatProvider? provider, out RangeRef result)
    {
        result = default;
        return s is not null && TryParse(s.AsSpan(), provider, out result);
    }

    /// <summary>Formats as <c>A1:C5</c>, or as <c>B2</c> for a single-cell range.</summary>
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
        if (!Start.TryFormat(destination, out int startLength, format, provider))
            return false;
        if (IsSingleCell)
        {
            charsWritten = startLength;
            return true;
        }

        if (destination.Length <= startLength)
            return false;
        destination[startLength] = ':';
        if (!End.TryFormat(destination[(startLength + 1)..], out int endLength, format, provider))
            return false;
        charsWritten = startLength + 1 + endLength;
        return true;
    }

    /// <inheritdoc/>
    public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
    {
        bytesWritten = 0;
        if (!Start.TryFormat(utf8Destination, out int startLength, format, provider))
            return false;
        if (IsSingleCell)
        {
            bytesWritten = startLength;
            return true;
        }

        if (utf8Destination.Length <= startLength)
            return false;
        utf8Destination[startLength] = (byte)':';
        if (!End.TryFormat(utf8Destination[(startLength + 1)..], out int endLength, format, provider))
            return false;
        bytesWritten = startLength + 1 + endLength;
        return true;
    }
}
