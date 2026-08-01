using System.Buffers;

namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Splits formula text into tokens without allocating: every token is a slice of the
/// source span. Whitespace is dropped except where it separates two operands, which is
/// Excel's intersection operator.
/// </summary>
internal ref struct FormulaLexer
{
    private static readonly SearchValues<char> ErrorChars = SearchValues.Create("#/!?_ABCDEFGHIJKLMNOPQRSTUVWXYZ0");

    private readonly ReadOnlySpan<char> _text;
    private int _position;
    private TokenKind _previous;

    public FormulaLexer(ReadOnlySpan<char> text)
    {
        _text = text;
        _position = 0;
        _previous = TokenKind.End;
    }

    /// <summary>Current offset in the source text.</summary>
    public readonly int Position => _position;

    /// <summary>Sheet name of the most recent <see cref="TokenKind.Reference"/> token, if any.</summary>
    public string? SheetName { get; private set; }

    /// <summary>Last sheet of a 3-D span prefix, or <see langword="null"/>.</summary>
    public string? SheetSpanEnd { get; private set; }

    /// <summary>One-based external book index of a <c>[N]Sheet!</c> prefix, or 0.</summary>
    public int ExternalBook { get; private set; }

    /// <summary>Parsed parts of the most recent <see cref="TokenKind.Reference"/> token.</summary>
    public RefPart Start { get; private set; }

    /// <summary>Whether the last reference token carried the spill suffix (<c>A1#</c>).</summary>
    public bool SpilledReference { get; private set; }

    /// <summary>Second part of the most recent reference token, for ranges.</summary>
    public RefPart End { get; private set; }

    /// <summary>Reads the next token.</summary>
    public FormulaToken Next()
    {
        bool sawSpace = SkipWhitespace();
        if (_position >= _text.Length)
            return Emit(TokenKind.End, _position, 0);

        if (sawSpace && CouldEndOperand(_previous) && CouldStartOperand(_text[_position]))
            return Emit(TokenKind.Space, _position, 0);

        char c = _text[_position];
        return c switch
        {
            '(' => Emit(TokenKind.OpenParen, _position, 1),
            ')' => Emit(TokenKind.CloseParen, _position, 1),
            '{' => Emit(TokenKind.OpenBrace, _position, 1),
            '}' => Emit(TokenKind.CloseBrace, _position, 1),
            ',' => Emit(TokenKind.Comma, _position, 1),
            ';' => Emit(TokenKind.Semicolon, _position, 1),
            '+' => Emit(TokenKind.Plus, _position, 1),
            '-' => Emit(TokenKind.Minus, _position, 1),
            '*' => Emit(TokenKind.Star, _position, 1),
            '/' => Emit(TokenKind.Slash, _position, 1),
            '^' => Emit(TokenKind.Caret, _position, 1),
            '%' => Emit(TokenKind.Percent, _position, 1),
            '&' => Emit(TokenKind.Ampersand, _position, 1),
            '=' => Emit(TokenKind.Equal, _position, 1),
            ':' => Emit(TokenKind.Colon, _position, 1),
            '@' => Emit(TokenKind.At, _position, 1),
            '<' => LessThan(),
            '>' => GreaterThan(),
            '"' => ScanString(),
            '#' => ScanError(),
            '[' => ScanBracket(),
            _ => ScanWord(c),
        };
    }

    /// <summary>
    /// A bracket opens either an external reference (<c>[1]Sheet1!A1</c>) or a structured
    /// reference in its bare <c>[@Column]</c> form; the book index is what tells them apart.
    /// </summary>
    private FormulaToken ScanBracket()
    {
        if (TryScanReference(out FormulaToken reference))
            return reference;
        return ScanTableReference(_position);
    }

    /// <summary>
    /// Reads a structured reference from its opening bracket to the matching close, keeping the
    /// nesting that <c>Table[[#Data],[Amount]]</c> needs. The body is parsed later: here it is
    /// enough to know where it ends.
    /// </summary>
    private FormulaToken ScanTableReference(int start)
    {
        int i = _position;
        int depth = 0;
        while (i < _text.Length)
        {
            char c = _text[i++];
            if (c == '\'' && i < _text.Length)
            {
                // Inside a column name an apostrophe escapes the next character.
                i++;
                continue;
            }

            if (c == '[')
            {
                depth++;
            }
            else if (c == ']' && --depth == 0)
            {
                return Emit(TokenKind.TableReference, start, i - start);
            }
        }

        throw new FormulaParseException("Unterminated structured reference", _text.ToString(), start);
    }

    private FormulaToken LessThan()
    {
        if (_position + 1 < _text.Length && _text[_position + 1] == '>')
            return Emit(TokenKind.NotEqual, _position, 2);
        if (_position + 1 < _text.Length && _text[_position + 1] == '=')
            return Emit(TokenKind.LessOrEqual, _position, 2);
        return Emit(TokenKind.Less, _position, 1);
    }

    private FormulaToken GreaterThan()
    {
        if (_position + 1 < _text.Length && _text[_position + 1] == '=')
            return Emit(TokenKind.GreaterOrEqual, _position, 2);
        return Emit(TokenKind.Greater, _position, 1);
    }

    private FormulaToken ScanString()
    {
        int start = _position;
        int i = _position + 1;
        bool escaped = false;
        while (i < _text.Length)
        {
            if (_text[i] != '"')
            {
                i++;
                continue;
            }

            if (i + 1 < _text.Length && _text[i + 1] == '"')
            {
                escaped = true;
                i += 2;
                continue;
            }

            break;
        }

        if (i >= _text.Length)
            throw new FormulaParseException("Unterminated string literal", _text.ToString(), start);

        ReadOnlySpan<char> body = _text[(start + 1)..i];
        _position = i + 1;
        _previous = TokenKind.String;
        // Doubled quotes are rare, so only escaped literals pay for materialization.
        return new FormulaToken(TokenKind.String, start, _position - start,
            escaped ? body.ToString().Replace("\"\"", "\"").AsSpan() : body);
    }

    private FormulaToken ScanError()
    {
        int start = _position;
        int i = _position;
        while (i < _text.Length && ErrorChars.Contains(char.ToUpperInvariant(_text[i])))
            i++;

        // "#N/A" and friends end in a letter; the scan above may also swallow a trailing
        // separator-free identifier, so trim back to the longest known error literal.
        while (i > start)
        {
            if (ExcelErrorText.TryParse(_text[start..i], out _))
                break;
            i--;
        }

        if (i == start)
            throw new FormulaParseException("Unknown error literal", _text.ToString(), start);

        return Emit(TokenKind.Error, start, i - start);
    }

    private FormulaToken ScanWord(char first)
    {
        int start = _position;
        if (char.IsAsciiDigit(first) || (first == '.' && _position + 1 < _text.Length && char.IsAsciiDigit(_text[_position + 1])))
        {
            if (TryScanReference(out FormulaToken rowReference))
                return rowReference;
            return ScanNumber();
        }

        if (TryScanReference(out FormulaToken reference))
            return reference;

        int i = _position;
        while (i < _text.Length && ReferenceScanner.IsNameChar(_text[i], i == start))
            i++;
        if (i == start)
            throw new FormulaParseException($"Unexpected character '{first}'", _text.ToString(), start);

        ReadOnlySpan<char> name = _text[start..i];
        if (name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return Emit(TokenKind.Boolean, start, i - start);

        return Emit(TokenKind.Name, start, i - start);
    }

    private FormulaToken ScanNumber()
    {
        int start = _position;
        int i = _position;
        while (i < _text.Length && char.IsAsciiDigit(_text[i]))
            i++;
        if (i < _text.Length && _text[i] == '.')
        {
            i++;
            while (i < _text.Length && char.IsAsciiDigit(_text[i]))
                i++;
        }

        if (i < _text.Length && (_text[i] is 'e' or 'E'))
        {
            int exponent = i + 1;
            if (exponent < _text.Length && _text[exponent] is '+' or '-')
                exponent++;
            if (exponent < _text.Length && char.IsAsciiDigit(_text[exponent]))
            {
                i = exponent;
                while (i < _text.Length && char.IsAsciiDigit(_text[i]))
                    i++;
            }
        }

        return Emit(TokenKind.Number, start, i - start);
    }

    /// <summary>
    /// Attempts to read a complete reference, including an optional sheet prefix and the
    /// range colon. Identifiers that merely look like cells (<c>LOG10</c> before a paren)
    /// stay names.
    /// </summary>
    private bool TryScanReference(out FormulaToken token)
    {
        token = default;
        int i = _position;
        int probe = i;
        string? spanEnd = null;
        int externalBook = 0;
        string? sheet =
            ReferenceScanner.TryScanSheetPrefix(_text, ref probe, out string parsedSheet, out spanEnd, out externalBook)
                ? parsedSheet
                : null;
        if (sheet is not null)
            i = probe;

        if (!ReferenceScanner.TryScanPart(_text, ref i, out RefPart start))
            return false;

        RefPart end = default;
        if (i < _text.Length && _text[i] == ':')
        {
            int afterColon = i + 1;
            if (ReferenceScanner.TryScanPart(_text, ref afterColon, out RefPart second) && second.Kind == start.Kind)
            {
                end = second;
                i = afterColon;
            }
        }

        // A bare column or row is only a reference as part of a range.
        if (start.Kind != RefPartKind.Cell && end.Kind == RefPartKind.None)
            return false;

        // An identifier followed by '(' is a function call, not a cell.
        if (sheet is null && end.Kind == RefPartKind.None && i < _text.Length && _text[i] == '(')
            return false;

        // A trailing '#' on a single cell is the spill suffix: the reference names whatever
        // the dynamic array at that cell spilled into. Only here does '#' not begin an error
        // literal — a range has no spill to name.
        bool spilled = false;
        if (start.Kind == RefPartKind.Cell && end.Kind == RefPartKind.None &&
            externalBook == 0 && i < _text.Length && _text[i] == '#')
        {
            spilled = true;
            i++;
        }

        SheetName = sheet;
        SheetSpanEnd = spanEnd;
        ExternalBook = externalBook;
        Start = start;
        End = end;
        SpilledReference = spilled;
        int length = i - _position;
        token = Emit(TokenKind.Reference, _position, length);
        return true;
    }

    private bool SkipWhitespace()
    {
        int start = _position;
        while (_position < _text.Length && char.IsWhiteSpace(_text[_position]))
            _position++;
        return _position > start;
    }

    private static bool CouldEndOperand(TokenKind kind) =>
        kind is TokenKind.Reference or TokenKind.Name or TokenKind.CloseParen or TokenKind.Number;

    private static bool CouldStartOperand(char c) =>
        char.IsAsciiLetterOrDigit(c) || c is '$' or '(' or '\'' or '_';

    private FormulaToken Emit(TokenKind kind, int start, int length)
    {
        _position = start + length;
        _previous = kind;
        return new FormulaToken(kind, start, length, _text.Slice(start, length));
    }
}
