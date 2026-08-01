namespace Cellwright.Formulas.Functions;

/// <summary>
/// <c>YEARFRAC</c> and the day-count conventions behind it. The five bases are the shared
/// vocabulary of every securities function, which is why they live apart from the date
/// catalog: COUPDAYS and PRICE count days with exactly these rules.
/// </summary>
internal static class DayCountFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("YEARFRAC", 2, 3, YearFrac);
    }

    private static FormulaValue YearFrac(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime start, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime end, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryInteger(context, args[2], out basis, out error))
            return FormulaValue.Scalar(error);

        if (basis is < 0 or > 4)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(YearFraction(start.Date, end.Date, basis));
    }

    /// <summary>
    /// The fraction of a year between two dates under the given basis (0 — US 30/360,
    /// 1 — actual/actual, 2 — actual/360, 3 — actual/365, 4 — European 30/360). The order of
    /// the dates does not matter; the answer is never negative.
    /// </summary>
    public static double YearFraction(DateTime start, DateTime end, int basis)
    {
        if (start > end)
            (start, end) = (end, start);

        return basis switch
        {
            1 => (end - start).Days / ActualYearLength(start, end),
            2 => (end - start).Days / 360.0,
            3 => (end - start).Days / 365.0,
            4 => Days360European(start, end) / 360.0,
            _ => Days360Nasd(start, end) / 360.0,
        };
    }

    /// <summary>
    /// The US (NASD) 30/360 day count: a 31st becomes the 30th — the end date only when the
    /// start already sits at a month's edge — and the end of February counts as the 30th.
    /// </summary>
    public static int Days360Nasd(DateTime start, DateTime end)
    {
        int startDay = start.Day;
        int endDay = end.Day;

        bool startAtFebruaryEnd = IsEndOfFebruary(start);
        if (startAtFebruaryEnd && IsEndOfFebruary(end))
            endDay = 30;
        if (startAtFebruaryEnd)
            startDay = 30;
        if (endDay == 31 && startDay >= 30)
            endDay = 30;
        if (startDay == 31)
            startDay = 30;

        return ((end.Year - start.Year) * 360) + ((end.Month - start.Month) * 30) + (endDay - startDay);
    }

    /// <summary>The European 30/360 day count: every 31st becomes the 30th, at either end.</summary>
    public static int Days360European(DateTime start, DateTime end)
    {
        int startDay = Math.Min(start.Day, 30);
        int endDay = Math.Min(end.Day, 30);
        return ((end.Year - start.Year) * 360) + ((end.Month - start.Month) * 30) + (endDay - startDay);
    }

    /// <summary>
    /// The denominator of an actual/actual fraction. Within one calendar year it is that
    /// year's length; across one year boundary but under a year apart, 366 exactly when a
    /// 29 February falls inside the range; further apart, the average length of the years
    /// crossed.
    /// </summary>
    private static double ActualYearLength(DateTime start, DateTime end)
    {
        if (start.Year == end.Year)
            return DateTime.IsLeapYear(start.Year) ? 366 : 365;

        bool withinOneYear = end.Year - start.Year == 1 &&
            (start.Month > end.Month || (start.Month == end.Month && start.Day >= end.Day));
        if (withinOneYear)
        {
            bool coversLeapDay =
                (DateTime.IsLeapYear(start.Year) && start <= new DateTime(start.Year, 2, 29)) ||
                (DateTime.IsLeapYear(end.Year) && end >= new DateTime(end.Year, 2, 29));
            return coversLeapDay ? 366 : 365;
        }

        int years = end.Year - start.Year + 1;
        double days = (new DateTime(end.Year + 1, 1, 1) - new DateTime(start.Year, 1, 1)).Days;
        return days / years;
    }

    private static bool IsEndOfFebruary(DateTime date) =>
        date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
}
