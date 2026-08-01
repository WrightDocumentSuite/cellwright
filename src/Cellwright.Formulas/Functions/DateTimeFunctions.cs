using System.Globalization;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// Date and time functions. Everything goes through the workbook's date system, so the 1904
/// option stays correct, and results are returned as real date-times rather than raw serials.
/// </summary>
internal static class DateTimeFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("TODAY", 0, 0, static (c, _) => FormulaValue.Scalar(CellValue.DateTime(c.Timestamp.Date)), isVolatile: true);
        builder.Add("NOW", 0, 0, static (c, _) => FormulaValue.Scalar(CellValue.DateTime(c.Timestamp)), isVolatile: true);
        builder.Add("DATE", 3, 3, Date);
        builder.Add("TIME", 3, 3, Time);
        builder.Add("YEAR", 1, 1, (c, a) => Part(c, a[0], static d => d.Year));
        builder.Add("MONTH", 1, 1, (c, a) => Part(c, a[0], static d => d.Month));
        builder.Add("DAY", 1, 1, (c, a) => Part(c, a[0], static d => d.Day));
        builder.Add("HOUR", 1, 1, (c, a) => Part(c, a[0], static d => d.Hour));
        builder.Add("MINUTE", 1, 1, (c, a) => Part(c, a[0], static d => d.Minute));
        builder.Add("SECOND", 1, 1, (c, a) => Part(c, a[0], static d => d.Second));
        builder.Add("WEEKDAY", 1, 2, Weekday);
        builder.Add("WEEKNUM", 1, 2, WeekNumber);
        builder.Add("EDATE", 2, 2, (c, a) => Shift(c, a, endOfMonth: false));
        builder.Add("EOMONTH", 2, 2, (c, a) => Shift(c, a, endOfMonth: true));
        builder.Add("DAYS", 2, 2, Days);
        builder.Add("DATEDIF", 3, 3, DateDif);
        builder.Add("DATEVALUE", 1, 1, DateValue);
        builder.Add("TIMEVALUE", 1, 1, TimeValue);
        builder.Add("WORKDAY", 2, 3, (c, a) => Workday(c, a, international: false), wantsRawReferences: true);
        builder.Add("NETWORKDAYS", 2, 3, (c, a) => NetworkDays(c, a, international: false), wantsRawReferences: true);
        builder.Add("WORKDAY.INTL", 2, 4, (c, a) => Workday(c, a, international: true), wantsRawReferences: true);
        builder.Add("NETWORKDAYS.INTL", 2, 4, (c, a) => NetworkDays(c, a, international: true), wantsRawReferences: true);
        builder.Add("DAYS360", 2, 3, Days360);
    }

    /// <summary>
    /// Counts days on a 360-day year of twelve 30-day months, the convention accounting uses to
    /// keep monthly interest equal. The two methods differ only in what they do with a 31st.
    /// </summary>
    /// <remarks>
    /// The American method moves the last day of February to the 30th as well as the 31st of a
    /// longer month, which the standard describes for the 31st alone (MS-OI29500 2.1.923). The
    /// European method moves only the 31st, at either end.
    /// </remarks>
    private static FormulaValue Days360(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryDate(context, args[0], out DateTime start, out CellValue error) ||
            !TryDate(context, args[1], out DateTime end, out error))
        {
            return FormulaValue.Scalar(error);
        }

        bool european = false;
        if (args.Length > 2 && !Args.TryBoolean(context, args[2], out european, out error))
            return FormulaValue.Scalar(error);

        int startDay = start.Day;
        int endDay = end.Day;

        if (startDay == 31 || (!european && IsEndOfFebruary(start)))
            startDay = 30;
        if (endDay == 31 && (european || startDay == 30))
            endDay = 30;

        return FormulaValue.Number(
            ((end.Year - start.Year) * 360) + ((end.Month - start.Month) * 30) + (endDay - startDay));
    }

    /// <summary>Whether a date is the last day of February, whichever length that February has.</summary>
    private static bool IsEndOfFebruary(DateTime date) =>
        date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);

    /// <summary>
    /// Moves a date by a number of working days, passing over the weekend and any holidays
    /// given. The <c>.INTL</c> form takes the week itself as an argument, ahead of the holidays.
    /// </summary>
    private static FormulaValue Workday(EvalContext context, ReadOnlySpan<FormulaValue> args, bool international)
    {
        if (!TryDate(context, args[0], out DateTime start, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int days, out error) ||
            !TryWeek(context, args, international, out Weekend weekend, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (weekend.Everything)
            return FormulaValue.Error(ExcelError.Value);

        HashSet<DateTime> holidays = CollectHolidays(context, args, international ? 3 : 2);
        DateTime current = start;
        int step = Math.Sign(days);

        // Each working day costs at most a full week of calendar days plus whatever holidays
        // fall in the way, so a walk longer than that is not going to arrive.
        long limit = ((long)Math.Abs(days) * 7) + holidays.Count + 7;
        for (int moved = 0; moved != days; limit--)
        {
            if (limit <= 0)
                return FormulaValue.Error(ExcelError.Number);

            current = current.AddDays(step);
            if (IsWorkday(current, weekend, holidays))
                moved += step;
        }

        return FormulaValue.Scalar(CellValue.DateTime(current.Date));
    }

    /// <summary>Counts working days between two dates, inclusive of both ends.</summary>
    private static FormulaValue NetworkDays(EvalContext context, ReadOnlySpan<FormulaValue> args, bool international)
    {
        if (!TryDate(context, args[0], out DateTime start, out CellValue error) ||
            !TryDate(context, args[1], out DateTime end, out error) ||
            !TryWeek(context, args, international, out Weekend weekend, out error))
        {
            return FormulaValue.Scalar(error);
        }

        HashSet<DateTime> holidays = CollectHolidays(context, args, international ? 3 : 2);
        int sign = start <= end ? 1 : -1;
        if (sign < 0)
            (start, end) = (end, start);

        int count = 0;
        for (DateTime day = start.Date; day <= end.Date; day = day.AddDays(1))
        {
            if (IsWorkday(day, weekend, holidays))
                count++;
        }

        return FormulaValue.Number(count * sign);
    }

    /// <summary>Reads the week the <c>.INTL</c> pair takes; the older pair works Monday to Friday.</summary>
    private static bool TryWeek(
        EvalContext context, ReadOnlySpan<FormulaValue> args, bool international, out Weekend weekend, out CellValue error)
    {
        if (international && args.Length > 2)
            return Weekend.TryParse(context, args[2], out weekend, out error);

        weekend = Weekend.SaturdaySunday;
        error = CellValue.Blank;
        return true;
    }

    private static bool IsWorkday(DateTime date, Weekend weekend, HashSet<DateTime> holidays) =>
        !weekend.Includes(date.DayOfWeek) && !holidays.Contains(date.Date);

    private static HashSet<DateTime> CollectHolidays(EvalContext context, ReadOnlySpan<FormulaValue> args, int index)
    {
        var holidays = new HashSet<DateTime>();
        if (args.Length <= index)
            return holidays;

        foreach (CellValue value in new ArgumentValues(args[index], context))
        {
            if (value.IsDateTime)
                holidays.Add(value.GetDateTime().Date);
            else if (value.IsNumber && value.GetNumber() >= 0)
                holidays.Add(ExcelSerialDate.ToDateTime(value.GetNumber(), context.DateSystem).Date);
        }

        return holidays;
    }

    private static FormulaValue Date(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int year, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int month, out error) ||
            !Args.TryInteger(context, args[2], out int day, out error))
        {
            return FormulaValue.Scalar(error);
        }

        // Excel maps two-digit years into 1900-1999 and rolls out-of-range months and days.
        if (year is >= 0 and < 1900)
            year += 1900;

        try
        {
            DateTime date = new DateTime(year, 1, 1).AddMonths(month - 1).AddDays(day - 1);
            return FormulaValue.Scalar(CellValue.DateTime(date));
        }
        catch (ArgumentOutOfRangeException)
        {
            return FormulaValue.Error(ExcelError.Number);
        }
    }

    private static FormulaValue Time(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int hours, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int minutes, out error) ||
            !Args.TryInteger(context, args[2], out int seconds, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double fraction = ((hours * 3600.0) + (minutes * 60.0) + seconds) / 86400.0;
        fraction -= Math.Floor(fraction);
        return FormulaValue.Scalar(CellValue.TimeSpan(TimeSpan.FromDays(fraction)));
    }

    private static FormulaValue Part(EvalContext context, in FormulaValue value, Func<DateTime, int> select)
    {
        if (!TryDate(context, value, out DateTime date, out CellValue error))
            return FormulaValue.Scalar(error);
        return FormulaValue.Number(select(date));
    }

    private static FormulaValue Weekday(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryDate(context, args[0], out DateTime date, out CellValue error))
            return FormulaValue.Scalar(error);

        int type = 1;
        if (args.Length > 1 && !Args.TryInteger(context, args[1], out type, out error))
            return FormulaValue.Scalar(error);

        int sunday = (int)date.DayOfWeek; // 0 = Sunday
        return type switch
        {
            1 => FormulaValue.Number(sunday + 1),
            2 => FormulaValue.Number(((sunday + 6) % 7) + 1),
            3 => FormulaValue.Number((sunday + 6) % 7),
            _ => FormulaValue.Error(ExcelError.Number),
        };
    }

    private static FormulaValue WeekNumber(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryDate(context, args[0], out DateTime date, out CellValue error))
            return FormulaValue.Scalar(error);

        var startOfYear = new DateTime(date.Year, 1, 1);
        int offset = (int)startOfYear.DayOfWeek;
        int week = ((date.DayOfYear + offset - 1) / 7) + 1;
        return FormulaValue.Number(week);
    }

    private static FormulaValue Shift(EvalContext context, ReadOnlySpan<FormulaValue> args, bool endOfMonth)
    {
        if (!TryDate(context, args[0], out DateTime date, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int months, out error))
        {
            return FormulaValue.Scalar(error);
        }

        try
        {
            DateTime shifted = date.AddMonths(months);
            if (endOfMonth)
                shifted = new DateTime(shifted.Year, shifted.Month, DateTime.DaysInMonth(shifted.Year, shifted.Month));
            return FormulaValue.Scalar(CellValue.DateTime(shifted.Date));
        }
        catch (ArgumentOutOfRangeException)
        {
            return FormulaValue.Error(ExcelError.Number);
        }
    }

    private static FormulaValue Days(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryDate(context, args[0], out DateTime end, out CellValue error) ||
            !TryDate(context, args[1], out DateTime start, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return FormulaValue.Number((end.Date - start.Date).TotalDays);
    }

    private static FormulaValue DateDif(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryDate(context, args[0], out DateTime start, out CellValue error) ||
            !TryDate(context, args[1], out DateTime end, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (start > end)
            return FormulaValue.Error(ExcelError.Number);

        string unit = Args.Text(context, args[2]).ToUpperInvariant();
        int months = ((end.Year - start.Year) * 12) + end.Month - start.Month - (end.Day < start.Day ? 1 : 0);
        return unit switch
        {
            "D" => FormulaValue.Number((end.Date - start.Date).TotalDays),
            "M" => FormulaValue.Number(months),
            "Y" => FormulaValue.Number(months / 12),
            "YM" => FormulaValue.Number(months % 12),
            "MD" => FormulaValue.Number((end.Date - start.Date.AddMonths(months)).TotalDays),
            "YD" => FormulaValue.Number((end.Date - start.Date.AddYears(months / 12)).TotalDays),
            _ => FormulaValue.Error(ExcelError.Number),
        };
    }

    private static FormulaValue DateValue(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)
            ? FormulaValue.Number(ExcelSerialDate.FromDateTime(date.Date, context.DateSystem))
            : FormulaValue.Error(ExcelError.Value);
    }

    private static FormulaValue TimeValue(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out DateTime date)
            ? FormulaValue.Number(date.TimeOfDay.TotalDays)
            : FormulaValue.Error(ExcelError.Value);
    }

    /// <summary>Reads an argument as a date, accepting serial numbers and date text.</summary>
    /// <summary>Shared with the day-count and securities functions, which all take dates.</summary>
    internal static bool TryDate(EvalContext context, in FormulaValue value, out DateTime date, out CellValue error)
    {
        date = default;
        CellValue scalar = Operators.Dereference(value, context);
        if (scalar.IsDateTime)
        {
            date = scalar.GetDateTime();
            error = CellValue.Blank;
            return true;
        }

        if (scalar.IsText && DateTime.TryParse(scalar.GetText(), CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
        {
            error = CellValue.Blank;
            return true;
        }

        if (!Args.TryNumber(context, value, out double serial, out error))
            return false;

        if (serial < 0)
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        try
        {
            date = ExcelSerialDate.ToDateTime(serial, context.DateSystem);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }
    }
}
