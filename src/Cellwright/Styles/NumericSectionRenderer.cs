using System.Globalization;
using System.Text;

namespace Cellwright.Styles;

/// <summary>
/// Lays a number out along a format section. The digits are produced once, up front, and then
/// spooled into the placeholders as they come: <c>#</c> takes a digit if there is one,
/// <c>0</c> takes one or a zero, and everything else is punctuation.
/// </summary>
internal static class NumericSectionRenderer
{
    /// <param name="section">One section of a format code.</param>
    /// <param name="number">Value to render, already absolute when the section owns the sign.</param>
    /// <param name="plan">What the section asked for.</param>
    /// <param name="culture">Culture for separators.</param>
    /// <param name="sectionSpellsOutSign">
    /// Whether the format has a negative section of its own, in which case the minus is that
    /// section's business rather than ours.
    /// </param>
    public static string Render(
        string section, double number, in FormatPlan plan, CultureInfo culture, bool sectionSpellsOutSign = false)
    {
        double scaled = Scale(number, plan);
        bool negative = scaled < 0;
        scaled = Math.Abs(scaled);

        if (plan.IsScientific)
            return RenderScientific(section, negative ? -scaled : scaled, plan, culture);

        string digits = scaled.ToString("F" + plan.DecimalDigits, CultureInfo.InvariantCulture);
        int point = digits.IndexOf('.', StringComparison.Ordinal);
        string whole = point < 0 ? digits : digits[..point];
        string fraction = point < 0 ? "" : digits[(point + 1)..];
        if (whole == "0" && plan.IntegerDigits == 0)
            whole = "";

        // A '#' after the point shows a decimal only when there is one, so the zeros rounding
        // produced beyond the mandatory places come back off.
        if (plan.DecimalDigits > plan.RequiredDecimals)
            fraction = fraction.TrimEnd('0').PadRight(plan.RequiredDecimals, '0');

        var builder = new StringBuilder();
        if (negative && !sectionSpellsOutSign)
            builder.Append(culture.NumberFormat.NegativeSign);

        Emit(builder, section, whole, fraction, plan, culture);
        return builder.ToString();
    }

    private static double Scale(double number, in FormatPlan plan)
    {
        for (int i = 0; i < plan.PercentScale; i++)
            number *= 100;
        for (int i = 0; i < plan.ThousandScale; i++)
            number /= 1000;
        return number;
    }

    private static void Emit(
        StringBuilder builder, string section, string whole, string fraction, in FormatPlan plan, CultureInfo culture)
    {
        int wholeSlots = CountIntegerSlots(section);

        // The digits sit at the right-hand end of the slots: with four slots and two digits the
        // first two are padding, and with more digits than slots the first one takes the rest.
        int padding = Math.Max(0, wholeSlots - whole.Length);
        int overflow = Math.Max(0, whole.Length - wholeSlots);
        int wholeIndex = 0;
        int fractionIndex = 0;
        int slotSeen = 0;
        bool afterPoint = false;

        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            switch (c)
            {
                case '"':
                    int end = NumberFormatRenderer.SkipQuoted(section, i);
                    builder.Append(section, i + 1, end - i - 1);
                    i = end;
                    break;

                case '[':
                    i = NumberFormatRenderer.SkipBracket(section, i);
                    break;

                case '\\' when i + 1 < section.Length:
                    builder.Append(section[++i]);
                    break;

                // An underscore reserves the width of the next character; an asterisk repeats
                // one to fill the cell. Neither has any width in a plain string.
                case '_' or '*':
                    i++;
                    break;

                case '0' or '#' or '?':
                    if (afterPoint)
                        EmitFractionDigit(builder, c, fraction, ref fractionIndex);
                    else
                        EmitWholeDigit(builder, c, whole, padding, overflow, ref wholeIndex, ref slotSeen, plan, culture);
                    break;

                case '.':
                    afterPoint = true;
                    if (fraction.Length > 0 || plan.RequiredDecimals > 0)
                        builder.Append(culture.NumberFormat.NumberDecimalSeparator);
                    break;

                // Between digits a comma asks for grouping and after them it scales; either way
                // it never reaches the output itself.
                case ',':
                    break;

                case '%':
                    builder.Append(culture.NumberFormat.PercentSymbol);
                    break;

                default:
                    builder.Append(c);
                    break;
            }
        }
    }

    /// <summary>
    /// Writes one integer placeholder: padding while the digits have not started, then digits,
    /// with the first digit slot taking everything that did not fit — which is how <c>#,##0</c>
    /// shows a million in full rather than truncating it.
    /// </summary>
    private static void EmitWholeDigit(
        StringBuilder builder, char placeholder, string whole, int padding, int overflow,
        ref int index, ref int slotSeen, in FormatPlan plan, CultureInfo culture)
    {
        int slot = slotSeen++;
        if (slot < padding)
        {
            if (placeholder == '0')
                builder.Append('0');
            else if (placeholder == '?')
                builder.Append(' ');
            return;
        }

        if (index >= whole.Length)
            return;

        int take = slot == padding ? overflow + 1 : 1;
        AppendGrouped(builder, whole.AsSpan(index, take), whole.Length - index, plan, culture);
        index += take;
    }

    /// <summary>Appends digits, inserting the group separator at the thousands boundaries.</summary>
    private static void AppendGrouped(
        StringBuilder builder, ReadOnlySpan<char> digits, int remaining, in FormatPlan plan, CultureInfo culture)
    {
        if (!plan.GroupThousands)
        {
            builder.Append(digits);
            return;
        }

        for (int i = 0; i < digits.Length; i++)
        {
            builder.Append(digits[i]);
            int left = remaining - i - 1;
            if (left > 0 && left % 3 == 0)
                builder.Append(culture.NumberFormat.NumberGroupSeparator);
        }
    }

    private static void EmitFractionDigit(StringBuilder builder, char placeholder, string fraction, ref int index)
    {
        if (index < fraction.Length)
        {
            builder.Append(fraction[index++]);
            return;
        }

        if (placeholder == '0')
            builder.Append('0');
        else if (placeholder == '?')
            builder.Append(' ');
    }

    private static int CountIntegerSlots(string section)
    {
        int slots = 0;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '"')
                i = NumberFormatRenderer.SkipQuoted(section, i);
            else if (c == '[')
                i = NumberFormatRenderer.SkipBracket(section, i);
            else if (c is '\\' or '_' or '*')
                i++;
            else if (c == '.')
                break;
            else if (c is '0' or '#' or '?')
                slots++;
        }

        return slots;
    }

    /// <summary>
    /// Scientific notation. The mantissa follows the digits before the <c>E</c> and the exponent
    /// the placeholders after it, so <c>0.00E+00</c> gives <c>1.23E+04</c>.
    /// </summary>
    private static string RenderScientific(string section, double number, in FormatPlan plan, CultureInfo culture)
    {
        int exponentDigits = CountExponentDigits(section);
        string text = number.ToString(
            "E" + Math.Max(plan.DecimalDigits, 0), CultureInfo.InvariantCulture);

        int e = text.IndexOf('E', StringComparison.Ordinal);
        string mantissa = text[..e].Replace(".", culture.NumberFormat.NumberDecimalSeparator, StringComparison.Ordinal);
        int exponent = int.Parse(text[(e + 1)..], CultureInfo.InvariantCulture);
        char sign = exponent < 0 ? '-' : '+';
        bool showsPlus = section.Contains("E+", StringComparison.OrdinalIgnoreCase);

        var builder = new StringBuilder(mantissa).Append('E');
        if (exponent < 0 || showsPlus)
            builder.Append(sign);
        return builder
            .Append(Math.Abs(exponent).ToString(CultureInfo.InvariantCulture).PadLeft(exponentDigits, '0'))
            .ToString();
    }

    private static int CountExponentDigits(string section)
    {
        int e = section.IndexOf('E', StringComparison.OrdinalIgnoreCase);
        if (e < 0)
            return 2;

        int digits = 0;
        for (int i = e + 1; i < section.Length; i++)
        {
            if (section[i] is '0' or '#' or '?')
                digits++;
            else if (section[i] is not ('+' or '-'))
                break;
        }

        return Math.Max(digits, 1);
    }
}
