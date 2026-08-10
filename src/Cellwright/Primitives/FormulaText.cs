using System.Globalization;
using System.Text;

namespace Cellwright;

/// <summary>
/// Textual surgery on the formulas the model stores as strings: renaming a sheet and moving
/// references when rows or columns are inserted or deleted.
/// </summary>
/// <remarks>
/// The formula engine lives in a separate package, so the core cannot compile a formula to
/// rewrite it. Working on the text instead keeps structural editing available to everyone and
/// keeps a round trip faithful: whatever the walker does not recognize is copied byte for byte,
/// so spacing, casing and exotic syntax survive untouched.
/// </remarks>
internal static class FormulaText
{
    /// <summary>
    /// Rewrites every <c>Sheet!</c> prefix naming <paramref name="oldName"/> so that it names
    /// <paramref name="newName"/>. Returns the original instance when nothing matched.
    /// </summary>
    public static string? RenameSheet(string? formula, string oldName, string newName)
    {
        if (string.IsNullOrEmpty(formula) || !formula.Contains('!'))
            return formula;

        var rewriter = new SheetRenamer(oldName, newName);
        return Rewrite(formula, ref rewriter);
    }

    /// <summary>
    /// Moves the references of a formula the way inserting or deleting rows or columns moves
    /// them. References into deleted cells become <c>#REF!</c>.
    /// </summary>
    /// <param name="formula">Formula text without the leading <c>=</c>.</param>
    /// <param name="shift">What moved, where, and on which sheet.</param>
    /// <param name="ownSheet">
    /// Name of the sheet the formula lives on, so that references written without a prefix are
    /// known to point at it.
    /// </param>
    public static string? Shift(string? formula, in ReferenceShift shift, string ownSheet)
    {
        if (string.IsNullOrEmpty(formula))
            return formula;

        var rewriter = new ReferenceShifter(shift, ownSheet);
        return Rewrite(formula, ref rewriter);
    }

    /// <summary>
    /// Copies a formula from one cell to another, moving only the relative axes of its A1
    /// references. This is the semantics Excel uses for formulas attached to a range, such as
    /// conditional-format conditions written relative to the range's top-left cell.
    /// </summary>
    public static string? Translate(string? formula, CellRef from, CellRef to)
    {
        if (string.IsNullOrEmpty(formula) || from == to)
            return formula;

        var rewriter = new RelativeReferenceTranslator(to.Row - from.Row, to.Column - from.Column);
        return Rewrite(formula, ref rewriter);
    }

    /// <summary>
    /// Renders a sheet name as a formula prefix, quoting it when Excel would: anything but
    /// letters, digits and underscores, or a leading digit, needs the apostrophes.
    /// </summary>
    public static string QuoteSheetName(string name)
    {
        if (!NeedsQuoting(name))
            return name;
        return $"'{name.Replace("'", "''", StringComparison.Ordinal)}'";
    }

    /// <summary>
    /// Renders a full prefix — book index, sheet, optional span end — quoted the way Excel
    /// quotes them: one pair of apostrophes around everything, bracket included.
    /// </summary>
    public static string QuotePrefix(int externalBook, string sheetName, string? spanEnd)
    {
        string book = externalBook > 0 ? $"[{externalBook}]" : "";
        string body = spanEnd is null ? sheetName : $"{sheetName}:{spanEnd}";
        if (!NeedsQuoting(sheetName) && (spanEnd is null || !NeedsQuoting(spanEnd)))
            return book + body;
        return $"'{(book + body).Replace("'", "''", StringComparison.Ordinal)}'";
    }

    private static bool NeedsQuoting(string name)
    {
        foreach (char c in name)
        {
            if (!char.IsAsciiLetterOrDigit(c) && c != '_' && c <= 127)
                return true;
        }

        return name.Length > 0 && char.IsAsciiDigit(name[0]);
    }

    /// <summary>
    /// Walks the text once, handing every A1 reference to the rewriter and copying everything
    /// else. String literals are skipped whole so that a range written inside one is left
    /// alone, and identifiers are consumed whole so that <c>LOG10</c> is never read as a cell.
    /// </summary>
    private static string Rewrite<TRewriter>(string formula, ref TRewriter rewriter)
        where TRewriter : struct, IReferenceRewriter
    {
        ReadOnlySpan<char> text = formula;
        StringBuilder? builder = null;
        int copied = 0;
        int position = 0;

        while (position < text.Length)
        {
            char c = text[position];
            switch (c)
            {
                case '"':
                    position = SkipStringLiteral(text, position);
                    continue;

                case '#':
                    position = SkipErrorLiteral(text, position);
                    continue;

                case '[' when !StartsSheetPrefix(text, position):
                    // A structured-reference column may itself look exactly like an A1 cell
                    // (Table1[A1], [@B2]). It is a name, not a coordinate to translate. An
                    // external-book prefix such as [1]Data! is deliberately left to the normal
                    // reference scanner below.
                    position = SkipBracketedReference(text, position);
                    continue;

                default:
                    break;
            }

            if (!CouldStartReference(c))
            {
                position++;
                continue;
            }

            int start = position;
            if (!TryScanReference(text, ref position, out ReferenceToken token))
            {
                position = SkipIdentifier(text, start);
                continue;
            }

            if (!rewriter.Rewrite(ref token))
                continue;

            builder ??= new StringBuilder(formula.Length + 8);
            builder.Append(text[copied..start]);
            token.AppendTo(builder);
            copied = position;
        }

        if (builder is null)
            return formula;

        builder.Append(text[copied..]);
        return builder.ToString();
    }

    /// <summary>
    /// Reads a complete reference: optional sheet prefix, one part, and for a range the colon
    /// and a matching second part. A bare column or row is only a reference inside a range,
    /// and an identifier in front of <c>(</c> is a function call.
    /// </summary>
    private static bool TryScanReference(ReadOnlySpan<char> text, ref int position, out ReferenceToken token)
    {
        token = default;
        int i = position;
        int probe = i;
        string? spanEnd = null;
        int externalBook = 0;
        string? sheet =
            ReferenceScanner.TryScanSheetPrefix(text, ref probe, out string parsed, out spanEnd, out externalBook)
                ? parsed
                : null;
        if (sheet is not null)
            i = probe;

        if (!ReferenceScanner.TryScanPart(text, ref i, out RefPart start))
            return false;

        RefPart end = default;
        if (i < text.Length && text[i] == ':')
        {
            int afterColon = i + 1;
            if (ReferenceScanner.TryScanPart(text, ref afterColon, out RefPart second) && second.Kind == start.Kind)
            {
                end = second;
                i = afterColon;
            }
        }

        if (start.Kind != RefPartKind.Cell && end.Kind == RefPartKind.None)
            return false;
        if (sheet is null && end.Kind == RefPartKind.None && i < text.Length && text[i] == '(')
            return false;

        token = new ReferenceToken
        {
            SheetName = sheet,
            SheetSpanEnd = spanEnd,
            ExternalBook = externalBook,
            Start = start,
            End = end,
        };
        position = i;
        return true;
    }

    private static int SkipStringLiteral(ReadOnlySpan<char> text, int position)
    {
        int i = position + 1;
        while (i < text.Length)
        {
            if (text[i] != '"')
            {
                i++;
                continue;
            }

            if (i + 1 < text.Length && text[i + 1] == '"')
            {
                i += 2;
                continue;
            }

            return i + 1;
        }

        return text.Length;
    }

    private static bool StartsSheetPrefix(ReadOnlySpan<char> text, int position)
    {
        int probe = position;
        return ReferenceScanner.TryScanSheetPrefix(
            text, ref probe, out _, out _, out _);
    }

    private static int SkipBracketedReference(ReadOnlySpan<char> text, int position)
    {
        int depth = 0;
        for (int i = position; i < text.Length; i++)
        {
            if (text[i] == '\'' && i + 1 < text.Length)
            {
                // Structured-reference column names use an apostrophe to escape the next
                // bracket (and other special characters).
                i++;
                continue;
            }

            if (text[i] == '[')
            {
                depth++;
            }
            else if (text[i] == ']' && --depth == 0)
            {
                return i + 1;
            }
        }

        return text.Length;
    }

    /// <summary>Skips <c>#REF!</c>, <c>#N/A</c> and friends, which would otherwise read as names.</summary>
    private static int SkipErrorLiteral(ReadOnlySpan<char> text, int position)
    {
        int i = position + 1;
        while (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '/' or '!' or '?' or '_'))
            i++;
        return i;
    }

    private static int SkipIdentifier(ReadOnlySpan<char> text, int position)
    {
        int i = position;
        while (i < text.Length && (ReferenceScanner.IsNameChar(text[i], first: false) || text[i] is '$' or '\''))
            i++;
        return Math.Max(i, position + 1);
    }

    private static bool CouldStartReference(char c) =>
        char.IsAsciiLetterOrDigit(c) || c is '$' or '_' or '\'' or '[' || c > 127;

    /// <summary>One A1 reference found in formula text, in the form a rewriter can edit.</summary>
    private struct ReferenceToken
    {
        /// <summary>Sheet the reference names, or <see langword="null"/> when it carries no prefix.</summary>
        public string? SheetName;

        /// <summary>Last sheet of a 3-D span, or <see langword="null"/>.</summary>
        public string? SheetSpanEnd;

        /// <summary>One-based external book index of a <c>[N]Sheet!</c> prefix, or 0.</summary>
        public int ExternalBook;

        /// <summary>Top-left part of the reference.</summary>
        public RefPart Start;

        /// <summary>Bottom-right part, or <see cref="RefPartKind.None"/> for a single cell.</summary>
        public RefPart End;

        /// <summary>Set by a rewriter when the reference no longer has a target.</summary>
        public bool IsBroken;

        public readonly void AppendTo(StringBuilder builder)
        {
            if (IsBroken)
            {
                builder.Append(ExcelError.Reference.ToText());
                return;
            }

            if (SheetName is { } sheet)
                builder.Append(QuotePrefix(ExternalBook, sheet, SheetSpanEnd)).Append('!');

            AppendPart(builder, Start);
            if (End.Kind != RefPartKind.None)
                AppendPart(builder.Append(':'), End);
        }

        private static void AppendPart(StringBuilder builder, RefPart part)
        {
            if (part.Kind is RefPartKind.Cell or RefPartKind.Column)
            {
                if (part.ColumnAbsolute)
                    builder.Append('$');
                builder.Append(ColumnName.Format(part.Column));
            }

            if (part.Kind is RefPartKind.Cell or RefPartKind.Row)
            {
                if (part.RowAbsolute)
                    builder.Append('$');
                builder.Append(part.Row.ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>Edits one reference; returns whether the text has to be replaced.</summary>
    private interface IReferenceRewriter
    {
        bool Rewrite(ref ReferenceToken token);
    }

    private readonly struct SheetRenamer(string oldName, string newName) : IReferenceRewriter
    {
        public bool Rewrite(ref ReferenceToken token)
        {
            // A sheet in another book keeps its name whatever this one renames its own to.
            if (token.ExternalBook > 0)
                return false;

            bool renamed = false;
            if (string.Equals(token.SheetName, oldName, StringComparison.OrdinalIgnoreCase))
            {
                token.SheetName = newName;
                renamed = true;
            }

            if (string.Equals(token.SheetSpanEnd, oldName, StringComparison.OrdinalIgnoreCase))
            {
                token.SheetSpanEnd = newName;
                renamed = true;
            }

            return renamed;
        }
    }

    private readonly struct ReferenceShifter(ReferenceShift shift, string ownSheet) : IReferenceRewriter
    {
        public bool Rewrite(ref ReferenceToken token)
        {
            // A 3-D span names one rectangle across many sheets with their own geometries;
            // no single shift is right for all of them, so the text stays as written. A cell
            // of another book is not this book's to move either.
            if (token.SheetSpanEnd is not null || token.ExternalBook > 0)
                return false;

            string sheet = token.SheetName ?? ownSheet;
            if (!string.Equals(sheet, shift.SheetName, StringComparison.OrdinalIgnoreCase))
                return false;

            return shift.Apply(ref token.Start, ref token.End, out token.IsBroken);
        }
    }

    /// <summary>Moves the unanchored row and column coordinates as a copied formula moves.</summary>
    private readonly struct RelativeReferenceTranslator(int rows, int columns) : IReferenceRewriter
    {
        public bool Rewrite(ref ReferenceToken token)
        {
            RefPart start = token.Start;
            RefPart end = token.End;
            bool changed = Translate(ref start, out bool broken);
            if (!broken && end.Kind != RefPartKind.None)
                changed |= Translate(ref end, out broken);

            if (!changed)
                return false;

            token.Start = start;
            token.End = end;
            token.IsBroken = broken;
            return true;
        }

        private bool Translate(ref RefPart part, out bool broken)
        {
            broken = false;
            int row = part.Row;
            int column = part.Column;

            if ((part.Kind is RefPartKind.Cell or RefPartKind.Row) && !part.RowAbsolute)
                row += rows;
            if ((part.Kind is RefPartKind.Cell or RefPartKind.Column) && !part.ColumnAbsolute)
                column += columns;

            if (((part.Kind is RefPartKind.Cell or RefPartKind.Row) && !SheetLimits.IsValidRow(row)) ||
                ((part.Kind is RefPartKind.Cell or RefPartKind.Column) && !SheetLimits.IsValidColumn(column)))
            {
                broken = true;
                return true;
            }

            if (row == part.Row && column == part.Column)
                return false;

            part = part with { Row = row, Column = column };
            return true;
        }
    }
}
