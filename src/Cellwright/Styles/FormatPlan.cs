namespace Cellwright.Styles;

/// <summary>
/// What one section of a number format asks for, worked out in a single scan before anything is
/// rendered. The digit counts decide the padding, and the scaling has to be known up front
/// because a trailing comma divides the number before any digit is placed.
/// </summary>
internal readonly record struct FormatPlan
{
    /// <summary>Whether the section spells out a date or a time rather than a number.</summary>
    public bool IsTemporal { get; init; }

    /// <summary>Whether the section shows an elapsed duration, written <c>[h]</c> or <c>[mm]</c>.</summary>
    public bool IsElapsed { get; init; }

    /// <summary>Whether a time in the section is on a twelve-hour clock.</summary>
    public bool HasMeridiem { get; init; }

    /// <summary>Mandatory digits before the decimal point, from the <c>0</c> placeholders.</summary>
    public int IntegerDigits { get; init; }

    /// <summary>Digit positions after the decimal point, mandatory and optional together.</summary>
    public int DecimalDigits { get; init; }

    /// <summary>Mandatory digits after the decimal point, from the <c>0</c> placeholders.</summary>
    public int RequiredDecimals { get; init; }

    /// <summary>Whether thousands are grouped, which a comma between digit placeholders asks for.</summary>
    public bool GroupThousands { get; init; }

    /// <summary>How many times the number is divided by a thousand, from the trailing commas.</summary>
    public int ThousandScale { get; init; }

    /// <summary>How many times the number is multiplied by a hundred, from the percent signs.</summary>
    public int PercentScale { get; init; }

    /// <summary>Whether the section uses scientific notation.</summary>
    public bool IsScientific { get; init; }

    /// <summary>
    /// Whether the section shows a ratio rather than decimals, as <c># ?/?</c> does. A date
    /// separates its parts with the same character, so this only holds where nothing temporal
    /// was seen.
    /// </summary>
    public bool IsFraction { get; init; }

    /// <summary>Reads a section and works out how it wants its number laid out.</summary>
    public static FormatPlan Analyze(string section)
    {
        bool temporal = false;
        bool elapsed = false;
        bool meridiem = false;
        bool scientific = false;
        bool group = false;
        int integerDigits = 0;
        int decimalDigits = 0;
        int requiredDecimals = 0;
        int percent = 0;
        int trailingCommas = 0;
        bool afterPoint = false;
        bool afterExponent = false;
        bool sawDigit = false;
        bool ratio = false;

        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            switch (c)
            {
                case '"':
                    i = NumberFormatRenderer.SkipQuoted(section, i);
                    break;

                case '\\' or '*' or '_':
                    i++;
                    break;

                case '[':
                    {
                        int end = NumberFormatRenderer.SkipBracket(section, i);
                        ReadOnlySpan<char> body = section.AsSpan(i + 1, end - i - 1);
                        if (!body.IsEmpty && char.ToLowerInvariant(body[0]) is 'h' or 'm' or 's')
                        {
                            temporal = true;
                            elapsed = true;
                        }

                        i = end;
                        break;
                    }

                case 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S':
                    temporal = true;
                    break;

                // A lone 'm' is a month or a minute depending on its neighbours, but either way
                // the section is temporal — unless it is the 'm' of "AM/PM", handled below.
                case 'm' or 'M':
                    temporal = true;
                    break;

                case 'A' or 'a':
                    if (Matches(section, i, "AM/PM") || Matches(section, i, "A/P"))
                    {
                        meridiem = true;
                        temporal = true;
                        i += Matches(section, i, "AM/PM") ? 4 : 2;
                    }

                    break;

                // The placeholders after the E belong to the exponent and say nothing about how
                // many decimals the mantissa carries.
                case '0' or '#' or '?' when afterExponent:
                    break;

                case '0' or '#' or '?':
                    sawDigit = true;
                    trailingCommas = 0;
                    if (afterPoint)
                    {
                        decimalDigits++;
                        if (c == '0')
                            requiredDecimals++;
                    }
                    else if (c == '0')
                    {
                        integerDigits++;
                    }

                    break;

                case '.':
                    afterPoint = true;
                    break;

                case ',':
                    if (sawDigit)
                        trailingCommas++;
                    group |= sawDigit && !afterPoint && HasDigitAfter(section, i);
                    break;

                case '%':
                    percent++;
                    break;

                // Placeholders either side of a slash ask for a ratio. A date uses the same
                // character, so this only counts where nothing has named a date part.
                case '/' when i > 0 && section[i - 1] is '0' or '#' or '?':
                    ratio |= i + 1 < section.Length &&
                        (section[i + 1] is '0' or '#' or '?' || char.IsAsciiDigit(section[i + 1]));
                    break;

                case 'E' or 'e':
                    if (i + 1 < section.Length && section[i + 1] is '+' or '-')
                    {
                        scientific = true;
                        afterExponent = true;
                    }

                    break;

                default:
                    break;
            }
        }

        return new FormatPlan
        {
            IsTemporal = temporal,
            IsElapsed = elapsed,
            HasMeridiem = meridiem,
            IntegerDigits = integerDigits,
            DecimalDigits = decimalDigits,
            RequiredDecimals = requiredDecimals,
            GroupThousands = group,
            ThousandScale = trailingCommas,
            PercentScale = percent,
            IsScientific = scientific,
            IsFraction = ratio && !temporal,
        };
    }

    private static bool HasDigitAfter(string section, int position)
    {
        for (int i = position + 1; i < section.Length; i++)
        {
            if (section[i] is '0' or '#' or '?')
                return true;
            if (section[i] != ',')
                return false;
        }

        return false;
    }

    private static bool Matches(string section, int position, string token) =>
        section.AsSpan(position).StartsWith(token, StringComparison.OrdinalIgnoreCase);
}
