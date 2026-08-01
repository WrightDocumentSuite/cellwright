namespace Cellwright.Formulas.Functions;

/// <summary>
/// The coupon calendar of a periodic security: dates step back from maturity by whole
/// periods, sticking to month ends when maturity sits on one. Everything the securities
/// functions know about days comes from here.
/// </summary>
internal readonly struct CouponSchedule
{
    private CouponSchedule(DateTime settlement, DateTime maturity, int frequency, int basis, DateTime previous, DateTime next, int count)
    {
        Settlement = settlement;
        Maturity = maturity;
        Frequency = frequency;
        Basis = basis;
        PreviousCoupon = previous;
        NextCoupon = next;
        Count = count;
    }

    public DateTime Settlement { get; }

    public DateTime Maturity { get; }

    public int Frequency { get; }

    public int Basis { get; }

    /// <summary>The last coupon date at or before settlement.</summary>
    public DateTime PreviousCoupon { get; }

    /// <summary>The first coupon date after settlement.</summary>
    public DateTime NextCoupon { get; }

    /// <summary>How many coupons remain from settlement to maturity — COUPNUM.</summary>
    public int Count { get; }

    /// <summary>Builds the calendar, refusing the argument shapes Excel refuses.</summary>
    public static bool TryCreate(DateTime settlement, DateTime maturity, int frequency, int basis, out CouponSchedule schedule)
    {
        schedule = default;
        if (settlement >= maturity || frequency is not (1 or 2 or 4) || basis is < 0 or > 4)
            return false;

        int step = 12 / frequency;
        DateTime next = maturity;
        int count = 0;
        while (true)
        {
            DateTime candidate = AddMonthsPinned(maturity, -(count + 1) * step);
            count++;
            if (candidate <= settlement)
            {
                schedule = new CouponSchedule(settlement, maturity, frequency, basis, candidate, next, count);
                return true;
            }

            next = candidate;
        }
    }

    /// <summary>
    /// Steps a date by months, keeping it on the month's last day when the anchor was on one —
    /// a coupon schedule set to 31 August pays at the end of February, not the 28th-and-broken.
    /// </summary>
    public static DateTime AddMonthsPinned(DateTime anchor, int months)
    {
        DateTime moved = anchor.AddMonths(months);
        bool anchorAtEnd = anchor.Day == DateTime.DaysInMonth(anchor.Year, anchor.Month);
        return anchorAtEnd
            ? new DateTime(moved.Year, moved.Month, DateTime.DaysInMonth(moved.Year, moved.Month))
            : moved;
    }

    /// <summary>Days in the coupon period holding settlement — COUPDAYS.</summary>
    public double PeriodLength => Basis switch
    {
        1 => (NextCoupon - PreviousCoupon).Days,
        3 => 365.0 / Frequency,
        _ => 360.0 / Frequency,
    };

    /// <summary>Days from the period's start to settlement — COUPDAYBS.</summary>
    public double DaysAccrued => Basis switch
    {
        0 => DayCountFunctions.Days360Nasd(PreviousCoupon, Settlement),
        4 => DayCountFunctions.Days360European(PreviousCoupon, Settlement),
        _ => (Settlement - PreviousCoupon).Days,
    };

    /// <summary>Days from settlement to the next coupon — COUPDAYSNC.</summary>
    public double DaysToNext => Basis switch
    {
        0 or 4 => PeriodLength - DaysAccrued,
        _ => (NextCoupon - Settlement).Days,
    };
}

/// <summary>The six COUP functions, all views of one <see cref="CouponSchedule"/>.</summary>
internal static class CouponFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("COUPPCD", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.DateTime(s.PreviousCoupon)));
        builder.Add("COUPNCD", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.DateTime(s.NextCoupon)));
        builder.Add("COUPNUM", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.Number(s.Count)));
        builder.Add("COUPDAYS", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.Number(s.PeriodLength)));
        builder.Add("COUPDAYBS", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.Number(s.DaysAccrued)));
        builder.Add("COUPDAYSNC", 3, 4, (c, a) => Coupon(c, a, static s => CellValue.Number(s.DaysToNext)));
    }

    private static FormulaValue Coupon(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Func<CouponSchedule, CellValue> read)
    {
        if (!TryReadSchedule(context, args, out CouponSchedule schedule, out CellValue error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Scalar(read(schedule));
    }

    /// <summary>Reads the shared argument shape: settlement, maturity, frequency, [basis].</summary>
    internal static bool TryReadSchedule(
        EvalContext context, ReadOnlySpan<FormulaValue> args, out CouponSchedule schedule, out CellValue error)
    {
        schedule = default;
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !Args.TryInteger(context, args[2], out int frequency, out error))
        {
            return false;
        }

        int basis = 0;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryInteger(context, args[3], out basis, out error))
            return false;

        if (!CouponSchedule.TryCreate(settlement.Date, maturity.Date, frequency, basis, out schedule))
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        error = CellValue.Blank;
        return true;
    }
}
