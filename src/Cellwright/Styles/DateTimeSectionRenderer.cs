using System.Globalization;
using System.Text;

namespace Cellwright.Styles;

/// <summary>
/// Renders a serial number through a date or time section. The one genuinely tricky part is
/// <c>m</c>: it means a month next to a year or a day, and minutes next to an hour or a second.
/// </summary>
internal static class DateTimeSectionRenderer
{
    public static string Render(
        string section, double serial, in FormatPlan plan, CultureInfo culture, ExcelDateSystem dateSystem)
    {
        var builder = new StringBuilder();
        TimeSpan elapsed = TimeSpan.FromDays(Math.Abs(serial));

        // An elapsed format has no calendar behind it, and a negative serial has no date at all;
        // the time of day still comes from the fractional part in both cases.
        DateTime moment = serial >= 0
            ? ExcelSerialDate.ToDateTime(serial, dateSystem)
            : DateTime.MinValue + TimeSpan.FromTicks(elapsed.Ticks % TimeSpan.TicksPerDay);
        int hour = plan.HasMeridiem ? TwelveHour(moment.Hour) : moment.Hour;

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
                    i = AppendElapsed(builder, section, i, elapsed);
                    break;

                case '\\' when i + 1 < section.Length:
                    builder.Append(section[++i]);
                    break;

                case '_' or '*':
                    i++;
                    break;

                case 'y' or 'Y':
                    i = AppendYear(builder, section, i, moment, culture);
                    break;

                case 'd' or 'D':
                    i = AppendDay(builder, section, i, moment, culture);
                    break;

                case 'h' or 'H':
                    i = AppendPadded(builder, section, i, hour);
                    break;

                case 's' or 'S':
                    i = AppendSeconds(builder, section, i, moment, culture);
                    break;

                case 'm' or 'M':
                    i = AppendMonthOrMinute(builder, section, i, moment, culture);
                    break;

                case 'A' or 'a':
                    int consumed = AppendMeridiem(builder, section, i, moment);
                    if (consumed == 0)
                        builder.Append(c);
                    i += consumed;
                    break;

                case '/':
                    builder.Append(culture.DateTimeFormat.DateSeparator);
                    break;

                default:
                    builder.Append(c);
                    break;
            }
        }

        return builder.ToString();
    }

    /// <summary>Writes an elapsed component, which is a total rather than a clock reading.</summary>
    private static int AppendElapsed(StringBuilder builder, string section, int position, TimeSpan elapsed)
    {
        int end = NumberFormatRenderer.SkipBracket(section, position);
        ReadOnlySpan<char> body = section.AsSpan(position + 1, end - position - 1);
        if (body.IsEmpty)
            return end;

        int width = body.Length;
        long total = char.ToLowerInvariant(body[0]) switch
        {
            'h' => (long)elapsed.TotalHours,
            'm' => (long)elapsed.TotalMinutes,
            's' => (long)elapsed.TotalSeconds,
            _ => long.MinValue,
        };

        if (total != long.MinValue)
            builder.Append(total.ToString(CultureInfo.InvariantCulture).PadLeft(width, '0'));
        return end;
    }

    private static int AppendYear(StringBuilder builder, string section, int position, DateTime moment, CultureInfo culture)
    {
        int count = Run(section, position);
        builder.Append(count <= 2
            ? (moment.Year % 100).ToString("00", culture)
            : moment.Year.ToString("0000", culture));
        return position + count - 1;
    }

    private static int AppendDay(StringBuilder builder, string section, int position, DateTime moment, CultureInfo culture)
    {
        int count = Run(section, position);
        builder.Append(count switch
        {
            1 => moment.Day.ToString(culture),
            2 => moment.Day.ToString("00", culture),
            3 => culture.DateTimeFormat.AbbreviatedDayNames[(int)moment.DayOfWeek],
            _ => culture.DateTimeFormat.DayNames[(int)moment.DayOfWeek],
        });
        return position + count - 1;
    }

    private static int AppendSeconds(StringBuilder builder, string section, int position, DateTime moment, CultureInfo culture)
    {
        int count = Run(section, position);
        builder.Append(count >= 2 ? moment.Second.ToString("00", culture) : moment.Second.ToString(culture));
        return position + count - 1;
    }

    /// <summary>
    /// Decides between month and minute. Excel reads the neighbours: an <c>m</c> that follows an
    /// hour or precedes seconds is minutes, and anything else is a month.
    /// </summary>
    private static int AppendMonthOrMinute(
        StringBuilder builder, string section, int position, DateTime moment, CultureInfo culture)
    {
        int count = Run(section, position);
        int end = position + count - 1;
        if (IsMinute(section, position, end))
        {
            builder.Append(count >= 2 ? moment.Minute.ToString("00", culture) : moment.Minute.ToString(culture));
            return end;
        }

        builder.Append(count switch
        {
            1 => moment.Month.ToString(culture),
            2 => moment.Month.ToString("00", culture),
            3 => culture.DateTimeFormat.AbbreviatedMonthNames[moment.Month - 1],
            4 => culture.DateTimeFormat.MonthNames[moment.Month - 1],
            _ => culture.DateTimeFormat.MonthNames[moment.Month - 1][..1],
        });
        return end;
    }

    private static bool IsMinute(string section, int start, int end)
    {
        for (int i = start - 1; i >= 0; i--)
        {
            char c = char.ToLowerInvariant(section[i]);
            if (c is 'h')
                return true;
            if (c is 'y' or 'd' or 's')
                return false;
            if (c is not (':' or ' ' or ']'))
                break;
            if (c == ']')
                return true;
        }

        for (int i = end + 1; i < section.Length; i++)
        {
            char c = char.ToLowerInvariant(section[i]);
            if (c is 's')
                return true;
            if (c is 'y' or 'd' or 'h')
                return false;
            if (c is not (':' or ' ' or '.'))
                break;
        }

        return false;
    }

    private static int AppendPadded(StringBuilder builder, string section, int position, int value)
    {
        int count = Run(section, position);
        builder.Append(count >= 2
            ? value.ToString("00", CultureInfo.InvariantCulture)
            : value.ToString(CultureInfo.InvariantCulture));
        return position + count - 1;
    }

    /// <summary>Writes the meridiem marker and returns how many extra characters it consumed.</summary>
    private static int AppendMeridiem(StringBuilder builder, string section, int position, DateTime moment)
    {
        ReadOnlySpan<char> rest = section.AsSpan(position);
        if (rest.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
        {
            builder.Append(moment.Hour < 12 ? "AM" : "PM");
            return "AM/PM".Length - 1;
        }

        if (rest.StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
        {
            builder.Append(moment.Hour < 12 ? 'A' : 'P');
            return "A/P".Length - 1;
        }

        return 0;
    }

    private static int TwelveHour(int hour) => hour % 12 == 0 ? 12 : hour % 12;

    /// <summary>Length of the run of the same letter starting at <paramref name="position"/>.</summary>
    private static int Run(string section, int position)
    {
        char first = char.ToLowerInvariant(section[position]);
        int count = 1;
        while (position + count < section.Length && char.ToLowerInvariant(section[position + count]) == first)
            count++;
        return count;
    }
}
