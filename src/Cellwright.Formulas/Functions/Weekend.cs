namespace Cellwright.Formulas.Functions;

/// <summary>
/// Which days of the week are not worked, as one bit per <see cref="DayOfWeek"/>.
/// </summary>
/// <remarks>
/// The plain <c>WORKDAY</c> and <c>NETWORKDAYS</c> assume Saturday and Sunday, which is a poor
/// assumption in much of the world; their <c>.INTL</c> pair takes the week as an argument
/// instead, and this is what that argument amounts to once it is read.
/// </remarks>
internal readonly record struct Weekend(int Days)
{
    /// <summary>The week the older pair of functions assumes.</summary>
    public static Weekend SaturdaySunday { get; } = new(Bit(DayOfWeek.Saturday) | Bit(DayOfWeek.Sunday));

    /// <summary>Whether no day of the week is worked, leaving a count of working days nothing to find.</summary>
    public bool Everything => Days == 0b111_1111;

    public bool Includes(DayOfWeek day) => (Days & Bit(day)) != 0;

    /// <summary>
    /// Reads the weekend argument: a code standing for one of the usual pairs or single days,
    /// or seven characters starting at Monday in which a one marks a day off.
    /// </summary>
    /// <returns>False when the argument names no week Excel recognizes.</returns>
    public static bool TryParse(
        EvalContext context, in FormulaValue value, out Weekend weekend, out CellValue error)
    {
        weekend = SaturdaySunday;
        error = CellValue.Blank;
        if (value.IsBlank)
            return true;

        CellValue scalar = Operators.Dereference(value, context);
        if (scalar.IsText)
            return TryPattern(scalar.GetText(), ref weekend, ref error);

        if (!Args.TryInteger(context, value, out int code, out error))
            return false;

        // Codes one to seven walk a pair of consecutive days round the week from Saturday;
        // eleven to seventeen name a single day, starting at Sunday.
        if (code is >= 1 and <= 7)
        {
            int first = ((int)DayOfWeek.Saturday + code - 1) % 7;
            weekend = new Weekend(Bit((DayOfWeek)first) | Bit((DayOfWeek)((first + 1) % 7)));
            return true;
        }

        if (code is >= 11 and <= 17)
        {
            weekend = new Weekend(Bit((DayOfWeek)(code - 11)));
            return true;
        }

        error = CellValue.Error(ExcelError.Number);
        return false;
    }

    private static bool TryPattern(string pattern, ref Weekend weekend, ref CellValue error)
    {
        if (pattern.Length != 7)
        {
            error = CellValue.Error(ExcelError.Value);
            return false;
        }

        int days = 0;
        for (int i = 0; i < 7; i++)
        {
            if (pattern[i] == '1')
                days |= Bit((DayOfWeek)((i + 1) % 7));   // The pattern starts at Monday.
            else if (pattern[i] != '0')
            {
                error = CellValue.Error(ExcelError.Value);
                return false;
            }
        }

        weekend = new Weekend(days);
        return true;
    }

    private static int Bit(DayOfWeek day) => 1 << (int)day;
}
