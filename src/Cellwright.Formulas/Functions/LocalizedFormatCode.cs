using System.Globalization;
using System.Text;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// Rewrites a number format string typed under a non-invariant locale into the invariant
/// placeholders the renderer understands. Excel stores the format argument of <c>TEXT</c>
/// verbatim, so a Russian file carries <c># ##0,00</c> — the comma is its decimal point and
/// the space its group separator, and reading either invariantly gets the number wrong.
/// </summary>
internal static class LocalizedFormatCode
{
    /// <summary>Digit placeholders a group separator has to sit between to mean grouping.</summary>
    private static bool IsDigitPlaceholder(char c) => c is '#' or '0' or '?';

    /// <summary>
    /// Translates the culture's decimal and group separators to <c>.</c> and <c>,</c>. Quoted
    /// literals, backslash escapes and bracketed sections travel untouched. Localized date
    /// letters are out of scope — only the separators change.
    /// </summary>
    public static string ToInvariant(string code, CultureInfo culture)
    {
        NumberFormatInfo numbers = culture.NumberFormat;
        string decimalSeparator = numbers.NumberDecimalSeparator;
        string groupSeparator = numbers.NumberGroupSeparator;
        if (decimalSeparator == "." && groupSeparator == ",")
            return code;

        // Locales whose group separator is a narrow or non-breaking space get it typed as a
        // plain space, which is also what the space in "# ##0" means to their Excel.
        bool spaceGroups = groupSeparator is "\u00A0" or "\u202F" or " ";

        var builder = new StringBuilder(code.Length);
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            switch (c)
            {
                case '"':
                    int closing = code.IndexOf('"', i + 1);
                    int end = closing < 0 ? code.Length - 1 : closing;
                    builder.Append(code, i, end - i + 1);
                    i = end;
                    continue;
                case '\\' when i + 1 < code.Length:
                    builder.Append(c).Append(code[i + 1]);
                    i++;
                    continue;
                case '[':
                    int bracket = code.IndexOf(']', i + 1);
                    int stop = bracket < 0 ? code.Length - 1 : bracket;
                    builder.Append(code, i, stop - i + 1);
                    i = stop;
                    continue;
                default:
                    break;
            }

            if (Matches(code, i, decimalSeparator))
            {
                builder.Append('.');
                i += decimalSeparator.Length - 1;
                continue;
            }

            bool isGroup = Matches(code, i, groupSeparator) || (spaceGroups && c == ' ');
            if (isGroup && Grouping(code, i, groupSeparator, c))
            {
                builder.Append(',');
                i += c == ' ' ? 0 : groupSeparator.Length - 1;
                continue;
            }

            builder.Append(c);
        }

        return builder.ToString();
    }

    private static bool Matches(string code, int index, string separator) =>
        separator.Length > 0 && code.AsSpan(index).StartsWith(separator, StringComparison.Ordinal);

    /// <summary>
    /// Whether the separator at <paramref name="index"/> sits between digit placeholders — the
    /// position where it means grouping. Anywhere else it is a literal: the space before the
    /// percent sign in <c>0,00 %</c> must stay a space.
    /// </summary>
    private static bool Grouping(string code, int index, string separator, char current)
    {
        if (index == 0)
            return false;

        int after = index + (current == ' ' ? 1 : separator.Length);
        return IsDigitPlaceholder(code[index - 1]) && after < code.Length && IsDigitPlaceholder(code[after]);
    }
}
