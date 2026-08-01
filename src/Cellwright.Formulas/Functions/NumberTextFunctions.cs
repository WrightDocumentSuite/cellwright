using System.Text;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// The functions that turn a number into a piece of text with a shape of its own: currency
/// (<c>DOLLAR</c>) and Roman numerals (<c>ROMAN</c>).
/// </summary>
internal static class NumberTextFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("DOLLAR", 1, 2, Dollar);
        builder.Add("USDOLLAR", 1, 2, Dollar);
        builder.Add("ROMAN", 1, 2, Roman);
        builder.Add("ARABIC", 1, 1, Arabic);
    }

    /// <summary>
    /// <c>DOLLAR(number, [decimals])</c> — the number as currency text, negatives in
    /// parentheses. The symbol is the dollar sign regardless of machine locale, like the rest
    /// of the culture-independent rendering.
    /// </summary>
    private static FormulaValue Dollar(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        int decimals = 2;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryInteger(context, args[1], out decimals, out error))
            return FormulaValue.Scalar(error);

        // Negative decimals round to the left of the point and print none.
        if (decimals < 0)
        {
            number = Args.Round(number, decimals);
            decimals = 0;
        }

        string digits = decimals > 0 ? "." + new string('0', decimals) : "";
        string code = $"$#,##0{digits};($#,##0{digits})";
        return FormulaValue.Text(Cellwright.Styles.NumberFormatRenderer.Format(
            CellValue.Number(number), code, System.Globalization.CultureInfo.InvariantCulture, context.DateSystem));
    }

    /// <summary>The letters of each form, largest first, subtractive pairs included.</summary>
    private static readonly (int Value, string Numeral)[][] RomanForms =
    [
        [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (950, "LM"), (900, "CM"), (500, "D"), (450, "LD"), (400, "CD"), (100, "C"), (95, "VC"), (90, "XC"), (50, "L"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"), (490, "XD"), (450, "LD"), (400, "CD"), (100, "C"), (99, "IC"), (95, "VC"), (90, "XC"), (50, "L"), (49, "IL"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (995, "VM"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"), (495, "VD"), (490, "XD"), (450, "LD"), (400, "CD"), (100, "C"), (99, "IC"), (95, "VC"), (90, "XC"), (50, "L"), (49, "IL"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (999, "IM"), (995, "VM"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"), (499, "ID"), (495, "VD"), (490, "XD"), (450, "LD"), (400, "CD"), (100, "C"), (99, "IC"), (95, "VC"), (90, "XC"), (50, "L"), (49, "IL"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
    ];

    /// <summary>
    /// <c>ROMAN(number, [form])</c> — forms 0 to 4 allow progressively bolder subtractions:
    /// classic <c>CDXCIX</c> shortens step by step to <c>ID</c>. TRUE means classic and FALSE
    /// simplified, as in Excel. Zero is empty text; anything outside 0..3999 is refused.
    /// </summary>
    private static FormulaValue Roman(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double raw, out CellValue error))
            return FormulaValue.Scalar(error);

        int form = 0;
        if (args.Length > 1 && !args[1].IsBlank)
        {
            CellValue selector = Operators.Dereference(args[1], context);
            if (selector.IsBoolean)
                form = selector.GetBoolean() ? 0 : 4;
            else if (Args.TryInteger(context, args[1], out form, out error) is false)
                return FormulaValue.Scalar(error);
        }

        int number = (int)Math.Truncate(raw);
        if (number is < 0 or > 3999 || form is < 0 or > 4)
            return FormulaValue.Error(ExcelError.Value);

        var text = new StringBuilder();
        foreach ((int value, string numeral) in RomanForms[form])
        {
            while (number >= value)
            {
                text.Append(numeral);
                number -= value;
            }
        }

        return FormulaValue.Text(text.ToString());
    }

    /// <summary>
    /// <c>ARABIC(text)</c> — the inverse, accepting any of the forms ROMAN produces: a letter
    /// standing before a larger one subtracts.
    /// </summary>
    private static FormulaValue Arabic(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]).Trim();
        if (text.Length == 0)
            return FormulaValue.Number(0);

        bool negative = text.StartsWith('-');
        if (negative)
            text = text[1..];

        int total = 0;
        int previous = 0;
        for (int i = text.Length - 1; i >= 0; i--)
        {
            int value = char.ToUpperInvariant(text[i]) switch
            {
                'I' => 1,
                'V' => 5,
                'X' => 10,
                'L' => 50,
                'C' => 100,
                'D' => 500,
                'M' => 1000,
                _ => 0,
            };
            if (value == 0)
                return FormulaValue.Error(ExcelError.Value);

            total += value < previous ? -value : value;
            previous = Math.Max(previous, value);
        }

        return FormulaValue.Number(negative ? -total : total);
    }
}
