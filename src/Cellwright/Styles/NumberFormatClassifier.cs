namespace Cellwright.Styles;

/// <summary>
/// How a number format renders numeric serial values.
/// </summary>
internal enum TemporalKind : byte
{
    /// <summary>A plain number.</summary>
    None = 0,

    /// <summary>A calendar date and/or time of day.</summary>
    Date = 1,

    /// <summary>An elapsed duration (formats like <c>[h]:mm:ss</c>).</summary>
    Duration = 2,
}

/// <summary>
/// Decides whether a number format identifier (builtin) or format code (custom)
/// represents a date, a duration or a plain number.
/// </summary>
internal static class NumberFormatClassifier
{
    /// <summary>Classifies by builtin id, or by the custom code when supplied.</summary>
    public static TemporalKind Classify(int numberFormatId, string? customFormatCode)
    {
        if (customFormatCode is not null)
            return ClassifyCode(customFormatCode);

        return numberFormatId switch
        {
            >= 14 and <= 22 => TemporalKind.Date,
            >= 27 and <= 36 => TemporalKind.Date,   // East Asian era formats.
            45 or 46 or 47 => TemporalKind.Duration,
            >= 50 and <= 58 => TemporalKind.Date,
            >= 71 and <= 81 => TemporalKind.Date,   // Thai locale formats.
            _ => TemporalKind.None,
        };
    }

    /// <summary>
    /// Classifies a custom format code by scanning for date/time tokens outside quoted
    /// literals, escapes and bracket sections. <c>[h]</c>/<c>[m]</c>/<c>[s]</c> mark durations.
    /// </summary>
    public static TemporalKind ClassifyCode(ReadOnlySpan<char> code)
    {
        bool sawDateToken = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            switch (c)
            {
                case '"':
                    int closing = code[(i + 1)..].IndexOf('"');
                    if (closing < 0)
                        return sawDateToken ? TemporalKind.Date : TemporalKind.None;
                    i += closing + 1;
                    break;

                case '\\':
                    i++;
                    break;

                case '[':
                    int end = code[(i + 1)..].IndexOf(']');
                    if (end < 0)
                        return sawDateToken ? TemporalKind.Date : TemporalKind.None;
                    ReadOnlySpan<char> section = code.Slice(i + 1, end);
                    if (!section.IsEmpty && section[0] is 'h' or 'H' or 'm' or 'M' or 's' or 'S')
                        return TemporalKind.Duration;
                    i += end + 1;
                    break;

                case 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M':
                    sawDateToken = true;
                    break;

                default:
                    break;
            }
        }

        return sawDateToken ? TemporalKind.Date : TemporalKind.None;
    }
}
