namespace Cellwright.Formulas;

/// <summary>
/// Thrown when a formula cannot be parsed. Malformed input never escapes as a
/// <see cref="StackOverflowException"/> or a generic failure: the parser reports the
/// offending position instead.
/// </summary>
public sealed class FormulaParseException : Exception
{
    /// <summary>Creates the exception for a position in the formula text.</summary>
    public FormulaParseException(string message, string formula, int position)
        : base($"{message} (position {position} in '{formula}')")
    {
        Formula = formula;
        Position = position;
    }

    /// <summary>The formula that failed to parse.</summary>
    public string Formula { get; }

    /// <summary>Zero-based character offset where parsing stopped.</summary>
    public int Position { get; }
}
