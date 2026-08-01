using System.Globalization;

namespace Cellwright;

/// <summary>Shape of one side of an A1 reference.</summary>
internal enum RefPartKind : byte
{
    /// <summary>Not a reference part.</summary>
    None,

    /// <summary>A full cell such as <c>B7</c>.</summary>
    Cell,

    /// <summary>Column letters only, valid inside <c>A:C</c>.</summary>
    Column,

    /// <summary>Row digits only, valid inside <c>1:5</c>.</summary>
    Row,
}

/// <summary>One parsed side of an A1 reference, with the anchors that survive shifting.</summary>
internal readonly record struct RefPart(RefPartKind Kind, int Row, int Column, bool RowAbsolute, bool ColumnAbsolute);

/// <summary>
/// Recognizes A1 references inside formula text. Scanning is deliberately strict: anything
/// that would fall outside the worksheet grid is left for the caller to treat as a name,
/// which is what Excel does with identifiers such as <c>XFE1</c>. Both the formula engine's
/// lexer and the core's textual rewriter read references through this one implementation.
/// </summary>
internal static class ReferenceScanner
{
    /// <summary>
    /// Tries to read one reference part starting at <paramref name="position"/>, advancing it
    /// past the part on success.
    /// </summary>
    public static bool TryScanPart(ReadOnlySpan<char> text, ref int position, out RefPart part)
    {
        part = default;
        int i = position;
        bool columnAbsolute = Eat(text, ref i, '$');

        int letterStart = i;
        while (i < text.Length && char.IsAsciiLetter(text[i]) && i - letterStart < 3)
            i++;
        int letterCount = i - letterStart;

        bool rowAbsolute = false;
        int digitStart;
        if (letterCount > 0)
        {
            int afterLetters = i;
            rowAbsolute = Eat(text, ref i, '$');
            digitStart = i;
            while (i < text.Length && char.IsAsciiDigit(text[i]))
                i++;

            if (i == digitStart)
            {
                // Letters alone form a column reference; a stray '$' belongs to nothing.
                if (rowAbsolute)
                    return false;
                if (!TryColumn(text[letterStart..afterLetters], out int columnOnly))
                    return false;
                part = new RefPart(RefPartKind.Column, 0, columnOnly, false, columnAbsolute);
                position = i;
                return true;
            }
        }
        else
        {
            digitStart = i;
            while (i < text.Length && char.IsAsciiDigit(text[i]))
                i++;
            if (i == digitStart)
                return false;

            // Digits alone form a row reference; the leading '$' anchored the row, not a column.
            if (!TryRow(text[digitStart..i], out int rowOnly))
                return false;
            part = new RefPart(RefPartKind.Row, rowOnly, 0, columnAbsolute, false);
            position = i;
            return true;
        }

        // A cell reference must not be glued to further identifier characters, so that
        // names such as A1B or LOG10X are not mistaken for cells.
        if (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.'))
            return false;

        if (!TryColumn(text[letterStart..(letterStart + letterCount)], out int column) ||
            !TryRow(text[digitStart..i], out int row))
        {
            return false;
        }

        part = new RefPart(RefPartKind.Cell, row, column, rowAbsolute, columnAbsolute);
        position = i;
        return true;
    }

    /// <summary>Parses a sheet name prefix ending in <c>!</c>, unquoting <c>'…'</c> form.</summary>
    public static bool TryScanSheetPrefix(ReadOnlySpan<char> text, ref int position, out string sheetName) =>
        TryScanSheetPrefix(text, ref position, out sheetName, out _, out _);

    /// <summary>
    /// Parses a sheet prefix in any of its shapes: <c>Sheet1!</c>, a 3-D span
    /// <c>Sheet1:Sheet3!</c>, the quoted forms of both, and an external book index
    /// <c>[2]Sheet1!</c>. A colon can only be the span separator — no sheet name may
    /// contain one — and the bracket number sits inside the quotes when there are any.
    /// </summary>
    /// <param name="text">Formula text.</param>
    /// <param name="position">Scan position, advanced past the <c>!</c> on success.</param>
    /// <param name="sheetName">First (or only) sheet of the prefix.</param>
    /// <param name="spanEnd">Last sheet of a 3-D span, or <see langword="null"/>.</param>
    /// <param name="externalBook">One-based external book index, or 0 for the own book.</param>
    public static bool TryScanSheetPrefix(
        ReadOnlySpan<char> text, ref int position, out string sheetName, out string? spanEnd, out int externalBook)
    {
        sheetName = "";
        spanEnd = null;
        externalBook = 0;
        int i = position;
        if (i < text.Length && text[i] == '\'')
        {
            i++;
            var builder = new System.Text.StringBuilder();
            while (i < text.Length)
            {
                if (text[i] == '\'')
                {
                    if (i + 1 < text.Length && text[i + 1] == '\'')
                    {
                        builder.Append('\'');
                        i += 2;
                        continue;
                    }

                    i++;
                    break;
                }

                builder.Append(text[i++]);
            }

            if (i >= text.Length || text[i] != '!')
                return false;

            string body = builder.ToString();
            if (!SplitPrefixBody(body, out sheetName, out spanEnd, out externalBook))
                return false;

            position = i + 1;
            return sheetName.Length > 0;
        }

        if (i < text.Length && text[i] == '[' && !TryScanBookIndex(text, ref i, out externalBook))
            return false;

        int start = i;
        while (i < text.Length && IsNameChar(text[i], i == start))
            i++;
        if (i == start)
            return false;

        sheetName = text[start..i].ToString();
        if (i < text.Length && text[i] == ':')
        {
            int afterColon = i + 1;
            int endStart = afterColon;
            while (afterColon < text.Length && IsNameChar(text[afterColon], afterColon == endStart))
                afterColon++;
            if (afterColon > endStart && afterColon < text.Length && text[afterColon] == '!')
            {
                spanEnd = text[endStart..afterColon].ToString();
                i = afterColon;
            }
        }

        if (i >= text.Length || text[i] != '!')
            return false;

        position = i + 1;
        return true;
    }

    /// <summary>Splits the unquoted body of a quoted prefix into its book, span and name parts.</summary>
    private static bool SplitPrefixBody(string body, out string sheetName, out string? spanEnd, out int externalBook)
    {
        sheetName = "";
        spanEnd = null;
        externalBook = 0;

        ReadOnlySpan<char> rest = body;
        if (rest.Length > 0 && rest[0] == '[')
        {
            int scan = 0;
            if (!TryScanBookIndex(rest, ref scan, out externalBook))
            {
                // Not an index: the bracket carries a whole path, as a decoded legacy formula
                // spells it — '[C:\Reports\Last year.xls]Sheet1'!A1. The body is one opaque
                // name; the colon of "C:" must not read as a span separator.
                sheetName = body;
                return true;
            }

            rest = rest[scan..];
        }

        int colon = rest.IndexOf(':');
        if (colon < 0)
        {
            sheetName = rest.ToString();
        }
        else
        {
            sheetName = rest[..colon].ToString();
            spanEnd = rest[(colon + 1)..].ToString();
            if (spanEnd.Length == 0)
                return false;
        }

        return sheetName.Length > 0;
    }

    /// <summary>Reads a <c>[N]</c> external book index, advancing past the closing bracket.</summary>
    private static bool TryScanBookIndex(ReadOnlySpan<char> text, ref int position, out int externalBook)
    {
        externalBook = 0;
        int i = position + 1;
        int digitStart = i;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
            i++;
        if (i == digitStart || i >= text.Length || text[i] != ']')
            return false;

        if (!int.TryParse(text[digitStart..i], NumberStyles.None, CultureInfo.InvariantCulture, out externalBook) ||
            externalBook <= 0)
        {
            return false;
        }

        position = i + 1;
        return true;
    }

    /// <summary>Whether the character can appear in a defined name or function name.</summary>
    public static bool IsNameChar(char c, bool first) =>
        char.IsAsciiLetter(c) || c == '_' || c == '\\' || c > 127 || (!first && (char.IsAsciiDigit(c) || c == '.' || c == '?'));

    private static bool Eat(ReadOnlySpan<char> text, ref int position, char c)
    {
        if (position < text.Length && text[position] == c)
        {
            position++;
            return true;
        }

        return false;
    }

    private static bool TryColumn(ReadOnlySpan<char> letters, out int column)
    {
        column = 0;
        foreach (char c in letters)
            column = column * 26 + ((c | 0x20) - 'a') + 1;
        return SheetLimits.IsValidColumn(column);
    }

    private static bool TryRow(ReadOnlySpan<char> digits, out int row) =>
        int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out row) && SheetLimits.IsValidRow(row);
}
