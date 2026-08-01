namespace Cellwright.Formulas.Parsing;

/// <summary>Which band of a table a structured reference names.</summary>
[Flags]
internal enum TableArea
{
    /// <summary>Nothing said; the body of the table, which is what <c>Table[Column]</c> means.</summary>
    None = 0,

    /// <summary>The rows between the headers and the totals.</summary>
    Data = 1,

    /// <summary>The header row.</summary>
    Headers = 2,

    /// <summary>The totals row.</summary>
    Totals = 4,

    /// <summary>Headers, data and totals together.</summary>
    All = Data | Headers | Totals,

    /// <summary>The single row the formula sits on, which is what <c>[@Column]</c> means.</summary>
    ThisRow = 8,
}

/// <summary>
/// A structured reference such as <c>Sales[Amount]</c> or <c>[@[Unit price]]</c>. Unlike an A1
/// reference it names a table rather than coordinates, so it resolves when the formula runs: a
/// table that grew a row moves every formula reading it without any of them being rewritten.
/// </summary>
/// <param name="TableName">Table named, or <see langword="null"/> for the table the formula sits in.</param>
/// <param name="FirstColumn">First column named, or <see langword="null"/> for every column.</param>
/// <param name="LastColumn">Last column of a span, or <see langword="null"/> when only one is named.</param>
/// <param name="Area">Band of the table the reference covers.</param>
/// <param name="Text">The reference as written, so that rendering it back is exact.</param>
internal readonly record struct TableRef(
    string? TableName, string? FirstColumn, string? LastColumn, TableArea Area, string Text)
{
    /// <summary>
    /// Parses a structured reference. The name in front of the brackets is optional: without it
    /// the reference means the table the formula lives in.
    /// </summary>
    /// <param name="tableName">Name written before the brackets, or <see langword="null"/>.</param>
    /// <param name="spec">The bracketed part, brackets included.</param>
    /// <param name="text">The whole reference as written.</param>
    public static TableRef Parse(string? tableName, ReadOnlySpan<char> spec, string text)
    {
        ReadOnlySpan<char> inner = spec.Length >= 2 ? spec[1..^1] : default;
        TableArea area = TableArea.None;
        string? first = null;
        string? last = null;
        bool span = false;

        int position = 0;
        while (position < inner.Length)
        {
            char c = inner[position];
            if (c is ',' or ' ')
            {
                position++;
                continue;
            }

            if (c == ':')
            {
                span = true;
                position++;
                continue;
            }

            if (c == '@')
            {
                area |= TableArea.ThisRow;
                position++;
                continue;
            }

            string item = ReadItem(inner, ref position);
            if (item.StartsWith('#'))
            {
                area |= AreaOf(item);
                continue;
            }

            if (span)
                last = item;
            else
                first ??= item;
        }

        return new TableRef(tableName, first, last, area, text);
    }

    /// <summary>Reads one item: a bracketed name or a bare run up to the next separator.</summary>
    private static string ReadItem(ReadOnlySpan<char> inner, ref int position)
    {
        if (inner[position] != '[')
        {
            int start = position;
            while (position < inner.Length && inner[position] is not (',' or ':'))
                position++;
            return Unescape(inner[start..position]);
        }

        position++;
        int bodyStart = position;
        while (position < inner.Length && inner[position] != ']')
        {
            // An apostrophe escapes the bracket or hash that follows it inside a column name.
            position += inner[position] == '\'' ? 2 : 1;
        }

        ReadOnlySpan<char> body = inner[bodyStart..Math.Min(position, inner.Length)];
        position = Math.Min(position + 1, inner.Length);
        return Unescape(body);
    }

    /// <summary>
    /// Drops the apostrophes that escape brackets and hashes inside a column name. Nothing else
    /// is touched: a column really can be called <c>"Decimal "</c>, trailing space and all, and
    /// tidying that away would leave the reference pointing at no column.
    /// </summary>
    private static string Unescape(ReadOnlySpan<char> text)
    {
        if (!text.Contains('\''))
            return text.ToString();

        var builder = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '\'' && i + 1 < text.Length)
                i++;
            builder.Append(text[i]);
        }

        return builder.ToString();
    }

    private static TableArea AreaOf(string item) => item switch
    {
        "#All" => TableArea.All,
        "#Headers" => TableArea.Headers,
        "#Totals" => TableArea.Totals,
        "#This Row" => TableArea.ThisRow,
        _ => TableArea.Data,
    };
}
