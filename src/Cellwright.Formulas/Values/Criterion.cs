namespace Cellwright.Formulas;

/// <summary>
/// A criterion of <c>SUMIF</c> and friends. Text criteria may embed a comparison operator
/// (<c>"&gt;=10"</c>), and plain text may use wildcards; everything else compares for equality.
/// </summary>
internal readonly struct Criterion
{
    private readonly CellValue _operand;
    private readonly Comparison _comparison;
    private readonly bool _wildcards;

    private Criterion(Comparison comparison, CellValue operand, bool wildcards)
    {
        _comparison = comparison;
        _operand = operand;
        _wildcards = wildcards;
    }

    private enum Comparison : byte
    {
        Equal,
        NotEqual,
        Less,
        LessOrEqual,
        Greater,
        GreaterOrEqual,
    }

    /// <summary>Interprets a criterion value.</summary>
    public static Criterion Parse(in CellValue value, EvalContext context)
    {
        if (!value.IsText)
            return new Criterion(Comparison.Equal, value, wildcards: false);

        ReadOnlySpan<char> text = value.GetText();
        Comparison comparison = Comparison.Equal;
        int skip = 0;
        if (text.StartsWith(">="))
        {
            comparison = Comparison.GreaterOrEqual;
            skip = 2;
        }
        else if (text.StartsWith("<="))
        {
            comparison = Comparison.LessOrEqual;
            skip = 2;
        }
        else if (text.StartsWith("<>"))
        {
            comparison = Comparison.NotEqual;
            skip = 2;
        }
        else if (text.StartsWith(">"))
        {
            comparison = Comparison.Greater;
            skip = 1;
        }
        else if (text.StartsWith("<"))
        {
            comparison = Comparison.Less;
            skip = 1;
        }
        else if (text.StartsWith("="))
        {
            skip = 1;
        }

        string operandText = text[skip..].ToString();
        CellValue operand = ParseOperand(operandText);
        bool wildcards = comparison is Comparison.Equal or Comparison.NotEqual && Wildcard.ContainsWildcards(operandText);
        return new Criterion(comparison, operand, wildcards);
    }

    /// <summary>Whether a value satisfies the criterion.</summary>
    public bool Matches(in CellValue value, EvalContext context)
    {
        if (_wildcards)
        {
            bool matched = value.IsText && Wildcard.IsMatch(value.GetText(), _operand.GetText());
            return _comparison == Comparison.NotEqual ? !matched : matched;
        }

        // An empty criterion counts blanks, and blanks never satisfy a numeric comparison.
        if (_operand.IsBlank || (_operand.IsText && _operand.GetText().Length == 0))
            return _comparison == Comparison.NotEqual ? !value.IsBlank : value.IsBlank;

        if (value.IsError)
            return false;

        int order = Coercion.Compare(value, _operand, context.DateSystem);
        return _comparison switch
        {
            Comparison.Equal => order == 0,
            Comparison.NotEqual => order != 0,
            Comparison.Less => order < 0,
            Comparison.LessOrEqual => order <= 0,
            Comparison.Greater => order > 0,
            _ => order >= 0,
        };
    }

    private static CellValue ParseOperand(string text)
    {
        if (text.Length == 0)
            return CellValue.Text("");
        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            return CellValue.Boolean(true);
        if (text.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return CellValue.Boolean(false);

        CellValue number = Coercion.ParseNumber(text);
        return number.IsError ? CellValue.Text(text) : number;
    }
}
