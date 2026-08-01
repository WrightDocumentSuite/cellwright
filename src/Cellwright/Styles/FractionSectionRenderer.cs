using System.Globalization;
using System.Text;

namespace Cellwright.Styles;

/// <summary>
/// Renders the fraction formats — <c># ?/?</c>, <c>?/???</c>, <c># ??/16</c> — which show a
/// number as a whole part and a ratio instead of as decimals.
/// </summary>
/// <remarks>
/// The room the format gives the two halves of the ratio is the whole specification:
/// <c>?/???</c> asks for the closest ratio whose denominator runs to three digits, while
/// <c># ??/16</c> fixes the denominator and only asks how many sixteenths there are. Finding the
/// closest ratio under a ceiling is what continued fractions are for, and the convergents they
/// produce are exactly the best approximations — with one gap, filled below.
/// </remarks>
internal static class FractionSectionRenderer
{
    public static string Render(string section, double number, CultureInfo culture, bool sectionSpellsOutSign)
    {
        if (!Layout.TryParse(section, out Layout layout))
            return number.ToString(culture);

        bool negative = number < 0;
        double magnitude = Math.Abs(number);
        long whole = layout.HasWholePart ? (long)Math.Floor(magnitude) : 0;
        (long numerator, long denominator) = Ratio(magnitude - whole, layout);

        // Rounding can push the ratio up to a whole one, which belongs in the whole part.
        if (denominator > 0 && numerator == denominator)
        {
            whole++;
            numerator = 0;
        }

        var builder = new StringBuilder();
        if (negative && !sectionSpellsOutSign)
            builder.Append(culture.NumberFormat.NegativeSign);

        AppendWhole(builder, layout, whole);
        AppendRatio(builder, layout, numerator, denominator);
        Spool(builder, layout.Suffix, "", blank: false);
        return builder.ToString();
    }

    private static (long Numerator, long Denominator) Ratio(double remainder, in Layout layout)
    {
        if (layout.FixedDenominator > 0)
        {
            long fixedDenominator = layout.FixedDenominator;
            return ((long)Math.Round(remainder * fixedDenominator, MidpointRounding.AwayFromZero), fixedDenominator);
        }

        return Approximate(remainder, Ceiling(layout.Denominator.Length));
    }

    private static long Ceiling(int digits) => digits switch
    {
        <= 0 => 9,
        >= 18 => long.MaxValue,
        _ => (long)Math.Pow(10, digits) - 1,
    };

    /// <summary>
    /// The closest ratio to <paramref name="value"/> whose denominator stays within the ceiling,
    /// by continued fractions.
    /// </summary>
    /// <remarks>
    /// The convergents are best approximations, but the closest ratio under a ceiling is
    /// sometimes not a convergent: it can be a part-way step towards the next one. When the next
    /// convergent overshoots the ceiling, the largest step towards it that still fits is tried
    /// as well, and whichever of the two lies nearer wins.
    /// </remarks>
    private static (long Numerator, long Denominator) Approximate(double value, long ceiling)
    {
        long previousNumerator = 0, previousDenominator = 1;
        long numerator = 1, denominator = 0;
        double remainder = value;

        for (int step = 0; step < 64; step++)
        {
            long term = (long)Math.Floor(remainder);
            long nextDenominator = (term * denominator) + previousDenominator;
            if (nextDenominator > ceiling || nextDenominator < 0)
                return Nearest(value, numerator, denominator, previousNumerator, previousDenominator, ceiling);

            long nextNumerator = (term * numerator) + previousNumerator;
            (previousNumerator, previousDenominator) = (numerator, denominator);
            (numerator, denominator) = (nextNumerator, nextDenominator);

            double fractional = remainder - term;
            if (fractional < 1e-12)
                break;

            remainder = 1 / fractional;
        }

        return denominator == 0 ? ((long)Math.Round(value), 1) : (numerator, denominator);
    }

    private static (long Numerator, long Denominator) Nearest(
        double value, long numerator, long denominator, long previousNumerator, long previousDenominator, long ceiling)
    {
        if (denominator == 0)
            return ((long)Math.Round(value), 1);

        long room = (ceiling - previousDenominator) / denominator;
        if (room <= 0)
            return (numerator, denominator);

        long candidateNumerator = (room * numerator) + previousNumerator;
        long candidateDenominator = (room * denominator) + previousDenominator;
        double convergent = Math.Abs(value - ((double)numerator / denominator));
        double candidate = Math.Abs(value - ((double)candidateNumerator / candidateDenominator));
        return candidate < convergent ? (candidateNumerator, candidateDenominator) : (numerator, denominator);
    }

    private static void AppendWhole(StringBuilder builder, in Layout layout, long whole)
    {
        // A whole part of nothing is shown only where the format demands a digit.
        string digits = whole == 0 && !layout.WholePart.Contains('0', StringComparison.Ordinal)
            ? ""
            : whole.ToString(CultureInfo.InvariantCulture);

        Spool(builder, layout.WholePart, digits, blank: false);
    }

    private static void AppendRatio(StringBuilder builder, in Layout layout, long numerator, long denominator)
    {
        // A number that came out whole has no ratio to show, and Excel leaves the room it would
        // have taken blank rather than writing a nought over something.
        bool blank = numerator == 0 && layout.HasWholePart &&
            !layout.Numerator.Contains('0', StringComparison.Ordinal);

        Spool(builder, layout.Numerator, blank ? "" : Digits(numerator), blank);
        if (!blank)
            builder.Append('/');
        else if (layout.Aligns)
            builder.Append(' ');

        string tail = blank || layout.FixedDenominator > 0 ? "" : Digits(denominator);
        Spool(builder, layout.Denominator, tail, blank);
    }

    private static string Digits(long value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Writes digits into a run of placeholders, right-aligned: the leftmost slot takes whatever
    /// did not fit, a <c>?</c> with nothing to show becomes a space and a <c>#</c> nothing at
    /// all. Anything that is not a placeholder is literal text and passes through.
    /// </summary>
    private static void Spool(StringBuilder builder, string format, string digits, bool blank)
    {
        int slots = 0;
        foreach (char c in format)
        {
            if (c is '0' or '#' or '?')
                slots++;
        }

        int padding = Math.Max(0, slots - digits.Length);
        int overflow = Math.Max(0, digits.Length - slots);
        int index = 0;
        int slot = 0;

        for (int i = 0; i < format.Length; i++)
        {
            char c = format[i];
            if (c == '"')
            {
                int end = NumberFormatRenderer.SkipQuoted(format, i);
                builder.Append(format, i + 1, end - i - 1);
                i = end;
                continue;
            }

            if (c == '[')
            {
                i = NumberFormatRenderer.SkipBracket(format, i);
                continue;
            }

            if (c is '\\' or '_' or '*')
            {
                if (c == '\\' && i + 1 < format.Length)
                    builder.Append(format[i + 1]);
                i++;
                continue;
            }

            if (c is not ('0' or '#' or '?'))
            {
                builder.Append(c);
                continue;
            }

            int current = slot++;
            if (blank || current < padding)
            {
                if (c == '0')
                    builder.Append('0');
                else if (c == '?')
                    builder.Append(' ');
                continue;
            }

            if (index >= digits.Length)
                continue;

            int take = current == padding ? overflow + 1 : 1;
            builder.Append(digits, index, take);
            index += take;
        }
    }

    /// <summary>How a fraction section divides up: whole part, numerator, denominator, the rest.</summary>
    private readonly record struct Layout(
        string WholePart, string Numerator, string Denominator, string Suffix, long FixedDenominator)
    {
        /// <summary>Whether the format sets aside room for a whole part before the ratio.</summary>
        public bool HasWholePart => WholePart.AsSpan().IndexOfAny("0#?") >= 0;

        /// <summary>Whether the ratio is padded to a fixed width, which a <c>?</c> asks for.</summary>
        public bool Aligns =>
            Numerator.Contains('?', StringComparison.Ordinal) || Denominator.Contains('?', StringComparison.Ordinal);

        public static bool TryParse(string section, out Layout layout)
        {
            layout = default;
            int slash = IndexOfSlash(section);
            if (slash <= 0 || slash + 1 >= section.Length)
                return false;

            int numeratorStart = slash;
            while (numeratorStart > 0 && section[numeratorStart - 1] is '0' or '#' or '?')
                numeratorStart--;
            if (numeratorStart == slash)
                return false;

            int end = slash + 1;
            while (end < section.Length && section[end] is '0' or '#' or '?')
                end++;

            long denominator = 0;
            if (end == slash + 1 && !TryFixed(section, slash, ref end, out denominator))
                return false;

            layout = new Layout(
                section[..numeratorStart],
                section[numeratorStart..slash],
                section[(slash + 1)..end],
                section[end..],
                denominator);
            return true;
        }

        /// <summary>Reads a denominator written out in full, as the sixteen of <c># ??/16</c>.</summary>
        private static bool TryFixed(string section, int slash, ref int end, out long denominator)
        {
            denominator = 0;
            while (end < section.Length && char.IsAsciiDigit(section[end]))
                end++;

            return end > slash + 1 &&
                long.TryParse(section.AsSpan(slash + 1, end - slash - 1), CultureInfo.InvariantCulture, out denominator) &&
                denominator > 0;
        }

        private static int IndexOfSlash(string section)
        {
            for (int i = 0; i < section.Length; i++)
            {
                switch (section[i])
                {
                    case '"':
                        i = NumberFormatRenderer.SkipQuoted(section, i);
                        break;
                    case '[':
                        i = NumberFormatRenderer.SkipBracket(section, i);
                        break;
                    case '\\':
                        i++;
                        break;
                    case '/':
                        return i;
                    default:
                        break;
                }
            }

            return -1;
        }
    }
}
