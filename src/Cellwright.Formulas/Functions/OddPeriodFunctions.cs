namespace Cellwright.Formulas.Functions;

/// <summary>
/// Bonds whose first or last coupon period is odd-sized (§18.17.7.238–241). The odd stretch is
/// measured in quasi-coupon periods — the regular grid continued through it — and each one
/// contributes its fraction to the accrued, counted and discounted day sums of the formulas.
/// </summary>
internal static class OddPeriodFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("ODDFPRICE", 8, 9, OddFirstPrice);
        builder.Add("ODDFYIELD", 8, 9, OddFirstYield);
        builder.Add("ODDLPRICE", 7, 8, (c, a) => OddLast(c, a, priceKnown: false));
        builder.Add("ODDLYIELD", 7, 8, (c, a) => OddLast(c, a, priceKnown: true));
    }

    /// <summary>The three day-fraction sums an odd stretch contributes.</summary>
    private readonly record struct OddStretch(double Counted, double Accrued, double ToCome);

    private static FormulaValue OddFirstPrice(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadOddFirst(context, args, out OddFirstBond bond, out CellValue error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(PriceOddFirst(bond, bond.SecondNumber));
    }

    private static FormulaValue OddFirstYield(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadOddFirst(context, args, out OddFirstBond bond, out CellValue error))
            return FormulaValue.Scalar(error);

        double price = bond.SecondNumber;
        if (price <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double low = (-0.9999) * bond.Schedule.Frequency;
        double high = 100;
        if (PriceOddFirst(bond, low) < price || PriceOddFirst(bond, high) > price)
            return FormulaValue.Error(ExcelError.Number);

        for (int i = 0; i < 200; i++)
        {
            double middle = (low + high) / 2;
            if (PriceOddFirst(bond, middle) > price)
                low = middle;
            else
                high = middle;
        }

        return FormulaValue.Number((low + high) / 2);
    }

    /// <summary>Everything ODDFPRICE and ODDFYIELD share once the arguments are read.</summary>
    private readonly record struct OddFirstBond(
        CouponSchedule Schedule, DateTime Issue, DateTime FirstCoupon, double Rate,
        double SecondNumber, double Redemption, int Basis);

    /// <summary>
    /// The price with an odd first period: the regular bond sum, except that the first coupon
    /// pays only the odd stretch's fraction and the accrued interest covers issue to settlement.
    /// </summary>
    private static double PriceOddFirst(in OddFirstBond bond, double yield)
    {
        CouponSchedule schedule = bond.Schedule;
        int frequency = schedule.Frequency;
        double coupon = 100 * bond.Rate / frequency;
        double periodYield = 1 + (yield / frequency);
        double toNext = schedule.DaysToNext / schedule.PeriodLength;
        int count = schedule.Count;

        OddStretch stretch = MeasureStretch(
            bond.Issue, schedule.Settlement, bond.Issue, bond.FirstCoupon, bond.FirstCoupon, frequency, bond.Basis);

        double price = bond.Redemption / Math.Pow(periodYield, count - 1 + toNext);
        price += coupon * stretch.Counted / Math.Pow(periodYield, toNext);
        for (int k = 2; k <= count; k++)
            price += coupon / Math.Pow(periodYield, k - 1 + toNext);
        return price - (coupon * stretch.Accrued);
    }

    private static bool TryReadOddFirst(
        EvalContext context, ReadOnlySpan<FormulaValue> args, out OddFirstBond bond, out CellValue error)
    {
        bond = default;
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !DateTimeFunctions.TryDate(context, args[2], out DateTime issue, out error) ||
            !DateTimeFunctions.TryDate(context, args[3], out DateTime firstCoupon, out error) ||
            !Args.TryNumber(context, args[4], out double rate, out error) ||
            !Args.TryNumber(context, args[5], out double second, out error) ||
            !Args.TryNumber(context, args[6], out double redemption, out error) ||
            !Args.TryInteger(context, args[7], out int frequency, out error))
        {
            return false;
        }

        int basis = 0;
        if (args.Length > 8 && !args[8].IsBlank && !Args.TryInteger(context, args[8], out basis, out error))
            return false;

        // The dates must run issue < settlement < first coupon < maturity, or the period is
        // not an odd first one.
        if (rate < 0 || redemption <= 0 ||
            !(issue.Date < settlement.Date && settlement.Date < firstCoupon.Date && firstCoupon.Date < maturity.Date) ||
            !CouponSchedule.TryCreate(settlement.Date, maturity.Date, frequency, basis, out CouponSchedule schedule))
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        bond = new OddFirstBond(schedule, issue.Date, firstCoupon.Date, rate, second, redemption, basis);
        error = CellValue.Blank;
        return true;
    }

    /// <summary>
    /// ODDLPRICE and ODDLYIELD: one closed formula each over the odd last stretch, which may
    /// span several quasi-coupon periods between the last interest date and maturity.
    /// </summary>
    private static FormulaValue OddLast(EvalContext context, ReadOnlySpan<FormulaValue> args, bool priceKnown)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !DateTimeFunctions.TryDate(context, args[2], out DateTime lastInterest, out error) ||
            !Args.TryNumber(context, args[3], out double rate, out error) ||
            !Args.TryNumber(context, args[4], out double second, out error) ||
            !Args.TryNumber(context, args[5], out double redemption, out error) ||
            !Args.TryInteger(context, args[6], out int frequency, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 7 && !args[7].IsBlank && !Args.TryInteger(context, args[7], out basis, out error))
            return FormulaValue.Scalar(error);

        settlement = settlement.Date;
        maturity = maturity.Date;
        lastInterest = lastInterest.Date;
        if (rate < 0 || redemption <= 0 || frequency is not (1 or 2 or 4) || basis is < 0 or > 4 ||
            !(lastInterest < settlement && settlement < maturity))
        {
            return FormulaValue.Error(ExcelError.Number);
        }

        // The quasi grid marches forward from the last interest date through maturity.
        OddStretch stretch = MeasureStretch(
            lastInterest, settlement, lastInterest, maturity, lastInterest, frequency, basis);

        double coupon = 100 * rate / frequency;
        if (!priceKnown)
        {
            double yield = second;
            double price = ((redemption + (stretch.Counted * coupon)) / (1 + (stretch.ToCome * yield / frequency))) -
                (stretch.Accrued * coupon);
            return FormulaValue.Number(price);
        }

        double known = second;
        if (known <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double numerator = redemption + (stretch.Counted * coupon) - (known + (stretch.Accrued * coupon));
        double yieldValue = numerator / (known + (stretch.Accrued * coupon)) * frequency / stretch.ToCome;
        return FormulaValue.Number(yieldValue);
    }

    /// <summary>
    /// Walks the quasi-coupon periods of an odd stretch and adds up the three day fractions:
    /// the stretch's own days (counted), the days accrued to settlement, and the days from
    /// settlement onward.
    /// </summary>
    private static OddStretch MeasureStretch(
        DateTime stretchStart, DateTime settlement, DateTime accrueFrom, DateTime stretchEnd,
        DateTime anchor, int frequency, int basis)
    {
        int step = 12 / frequency;

        DateTime periodEnd = anchor;
        while (CouponSchedule.AddMonthsPinned(periodEnd, -step) > stretchStart)
            periodEnd = CouponSchedule.AddMonthsPinned(periodEnd, -step);
        if (periodEnd <= stretchStart)
            periodEnd = CouponSchedule.AddMonthsPinned(periodEnd, step);

        double counted = 0;
        double accrued = 0;
        double toCome = 0;
        DateTime periodStart = CouponSchedule.AddMonthsPinned(periodEnd, -step);
        while (periodStart < stretchEnd)
        {
            double length = basis switch
            {
                1 => (periodEnd - periodStart).Days,
                3 => 365.0 / frequency,
                _ => 360.0 / frequency,
            };

            counted += Overlap(periodStart, periodEnd, stretchStart, stretchEnd, basis) / length;
            accrued += Overlap(periodStart, periodEnd, accrueFrom, settlement, basis) / length;
            toCome += Overlap(periodStart, periodEnd, settlement, stretchEnd, basis) / length;

            periodStart = periodEnd;
            periodEnd = CouponSchedule.AddMonthsPinned(periodEnd, step);
        }

        return new OddStretch(counted, accrued, toCome);
    }

    /// <summary>Basis-days of the overlap between a quasi period and a stretch of interest.</summary>
    private static double Overlap(
        DateTime periodStart, DateTime periodEnd, DateTime from, DateTime to, int basis)
    {
        DateTime start = periodStart > from ? periodStart : from;
        DateTime end = periodEnd < to ? periodEnd : to;
        if (end <= start)
            return 0;

        return basis switch
        {
            0 => DayCountFunctions.Days360Nasd(start, end),
            4 => DayCountFunctions.Days360European(start, end),
            _ => (end - start).Days,
        };
    }
}

/// <summary>
/// The functions whose answers live outside the file: cube queries need a data-model
/// connection, RTD a running server, GETPIVOTDATA a pivot table. None of those exists in this
/// engine, and the honest result is the error Excel itself shows when the source is gone —
/// not <c>#NAME?</c>, which would say the function is unknown.
/// </summary>
internal static class ConnectionStubFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("CUBEVALUE", 1, FunctionDefinition.Unlimited, NotAvailable);
        builder.Add("CUBEMEMBER", 2, 3, NotAvailable);
        builder.Add("CUBEMEMBERPROPERTY", 3, 3, NotAvailable);
        builder.Add("CUBERANKEDMEMBER", 3, 4, NotAvailable);
        builder.Add("CUBESET", 2, 5, NotAvailable);
        builder.Add("CUBESETCOUNT", 1, 1, NotAvailable);
        builder.Add("CUBEKPIMEMBER", 3, 4, NotAvailable);
        builder.Add("RTD", 2, FunctionDefinition.Unlimited, NotAvailable);
        builder.Add("GETPIVOTDATA", 2, FunctionDefinition.Unlimited, Reference, wantsRawReferences: true);
    }

    private static FormulaValue NotAvailable(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        FormulaValue.Error(ExcelError.NotAvailable);

    private static FormulaValue Reference(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        FormulaValue.Error(ExcelError.Reference);
}
