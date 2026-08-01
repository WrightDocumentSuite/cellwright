namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Lexical categories of a formula.
/// </summary>
internal enum TokenKind : byte
{
    End,
    Number,
    String,
    Boolean,
    Error,
    /// <summary>An A1 cell or range reference, optionally sheet-qualified.</summary>
    Reference,
    /// <summary>An identifier: a function name (when followed by <c>(</c>) or a defined name.</summary>
    Name,
    /// <summary>A structured table reference such as <c>Sales[Amount]</c> or <c>[@Price]</c>.</summary>
    TableReference,
    OpenParen,
    CloseParen,
    OpenBrace,
    CloseBrace,
    Comma,
    Semicolon,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Percent,
    Ampersand,
    Equal,
    NotEqual,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,
    Colon,
    /// <summary>Whitespace between references: the intersection operator.</summary>
    Space,
    /// <summary>The <c>@</c> of a modern formula, which narrows a range to one value.</summary>
    At,
}

/// <summary>
/// One token: its kind, the slice of source it covers and the decoded payload for literals.
/// </summary>
internal readonly ref struct FormulaToken
{
    public FormulaToken(TokenKind kind, int start, int length, ReadOnlySpan<char> text)
    {
        Kind = kind;
        Start = start;
        Length = length;
        Text = text;
    }

    public TokenKind Kind { get; }

    /// <summary>Offset of the token in the formula text.</summary>
    public int Start { get; }

    /// <summary>Length of the token in the formula text.</summary>
    public int Length { get; }

    /// <summary>
    /// Payload of the token: the literal contents for strings (unescaped), the raw text
    /// for numbers, references and names.
    /// </summary>
    public ReadOnlySpan<char> Text { get; }
}
