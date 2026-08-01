using System.Globalization;
using System.Text;

namespace Cellwright.Styles;

/// <summary>
/// Renders a value the way a number format code would display it: <c>#,##0.00</c>,
/// <c>yyyy-mm-dd</c>, <c>0.00%</c> and the rest of the vocabulary Excel puts in a cell's format
/// and in the second argument of <c>TEXT</c>.
/// </summary>
/// <remarks>
/// A format code has up to four sections separated by <c>;</c>: positive, negative, zero and
/// text. With fewer sections Excel reuses them by the documented rules, which is what
/// <see cref="SelectSection"/> implements. Which of the three renderers a section goes to is
/// decided by <see cref="FormatPlan"/> in one scan: a date, a ratio, or plain digits.
/// </remarks>
internal static class NumberFormatRenderer
{
    /// <summary>Formats a value with a number format code.</summary>
    /// <param name="value">Value to render; blank renders as an empty string.</param>
    /// <param name="code">Format code, e.g. <c>#,##0.00;[Red]-#,##0.00</c>.</param>
    /// <param name="culture">Culture for separators and month names.</param>
    /// <param name="dateSystem">Date system for turning serial numbers into dates.</param>
    public static string Format(
        in CellValue value, string code, CultureInfo culture, ExcelDateSystem dateSystem = ExcelDateSystem.Date1900)
    {
        if (string.IsNullOrEmpty(code) || code.Equals("General", StringComparison.OrdinalIgnoreCase))
            return General(value, culture);

        List<string> sections = Split(code);
        if (value.IsText || value.IsBoolean || value.IsError)
        {
            // Only a fourth section applies to text; without one the text passes through.
            string text = value.ToString();
            return sections.Count >= 4 ? RenderText(sections[3], text) : text;
        }

        if (value.IsBlank)
            return "";

        double number = value.ToSerialNumber(dateSystem);
        string section = SelectSection(sections, number);

        // A format with its own negative section renders the absolute value: whatever sign the
        // reader is meant to see, parentheses or a minus, is written in that section.
        bool ownSign = sections.Count >= 2 && number < 0;
        return Render(section, ownSign ? Math.Abs(number) : number, culture, dateSystem, ownSign);
    }

    /// <summary>What Excel shows for a value with no format of its own.</summary>
    private static string General(in CellValue value, CultureInfo culture) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Number => value.GetNumber().ToString("G15", culture),
        _ => value.ToString(),
    };

    /// <summary>Splits on the section separator, ignoring semicolons inside literals and brackets.</summary>
    private static List<string> Split(string code)
    {
        List<string> sections = [];
        int start = 0;
        for (int i = 0; i < code.Length; i++)
        {
            switch (code[i])
            {
                case '"':
                    i = SkipQuoted(code, i);
                    break;
                case '[':
                    i = SkipBracket(code, i);
                    break;
                case '\\':
                    i++;
                    break;
                case ';':
                    sections.Add(code[start..i]);
                    start = i + 1;
                    break;
                default:
                    break;
            }
        }

        sections.Add(code[start..]);
        return sections;
    }

    /// <summary>
    /// Picks the section for a number. One section covers everything; two split at zero, with
    /// the negative one rendering the absolute value; three add a separate zero section.
    /// </summary>
    private static string SelectSection(List<string> sections, double number) => sections.Count switch
    {
        1 => sections[0],
        2 => number < 0 ? sections[1] : sections[0],
        _ => number < 0 ? sections[1] : number == 0 ? sections[2] : sections[0],
    };

    private static string RenderText(string section, string text)
    {
        var builder = new StringBuilder();
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            switch (c)
            {
                case '@':
                    builder.Append(text);
                    break;
                case '"':
                    int end = SkipQuoted(section, i);
                    builder.Append(section, i + 1, end - i - 1);
                    i = end;
                    break;
                case '[':
                    i = SkipBracket(section, i);
                    break;
                case '\\' when i + 1 < section.Length:
                    builder.Append(section[++i]);
                    break;
                case '*' or '_' when i + 1 < section.Length:
                    i++;
                    break;
                default:
                    builder.Append(c);
                    break;
            }
        }

        return builder.ToString();
    }

    private static string Render(
        string section, double number, CultureInfo culture, ExcelDateSystem dateSystem, bool sectionSpellsOutSign)
    {
        var plan = FormatPlan.Analyze(section);
        if (plan.IsTemporal)
            return DateTimeSectionRenderer.Render(section, number, plan, culture, dateSystem);

        return plan.IsFraction
            ? FractionSectionRenderer.Render(section, number, culture, sectionSpellsOutSign)
            : NumericSectionRenderer.Render(section, number, plan, culture, sectionSpellsOutSign);
    }

    /// <summary>Index of the closing quote of a literal starting at <paramref name="position"/>.</summary>
    internal static int SkipQuoted(string code, int position)
    {
        int end = code.IndexOf('"', position + 1);
        return end < 0 ? code.Length - 1 : end;
    }

    /// <summary>Index of the closing bracket of a section starting at <paramref name="position"/>.</summary>
    internal static int SkipBracket(string code, int position)
    {
        int end = code.IndexOf(']', position + 1);
        return end < 0 ? code.Length - 1 : end;
    }
}
