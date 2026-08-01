namespace Cellwright.Formulas.Functions;

/// <summary>
/// Securities priced per 100 of face value: the coupon bond pair PRICE and YIELD, the
/// discounted family, bills, accrued interest and Macaulay's duration. The day arithmetic all
/// comes from <see cref="CouponSchedule"/> and <see cref="DayCountFunctions"/>.
/// </summary>
internal static class SecurityFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("PRICE", 6, 7, Price);
        builder.Add("YIELD", 6, 7, Yield);
        builder.Add("DURATION", 5, 6, (c, a) => Duration(c, a, modified: false));
        builder.Add("MDURATION", 5, 6, (c, a) => Duration(c, a, modified: true));
        builder.Add("DISC", 4, 5, Disc);
        builder.Add("PRICEDISC", 4, 5, PriceDisc);
        builder.Add("YIELDDISC", 4, 5, YieldDisc);
        builder.Add("INTRATE", 4, 5, IntRate);
        builder.Add("RECEIVED", 4, 5, Received);
        builder.Add("PRICEMAT", 5, 6, PriceMat);
        builder.Add("YIELDMAT", 5, 6, YieldMat);
        builder.Add("TBILLPRICE", 3, 3, (c, a) => Bill(c, a, BillKind.Price));
        builder.Add("TBILLYIELD", 3, 3, (c, a) => Bill(c, a, BillKind.Yield));
        builder.Add("TBILLEQ", 3, 3, (c, a) => Bill(c, a, BillKind.Equivalent));
        builder.Add("ACCRINT", 6, 8, AccruedPeriodic);
        builder.Add("ACCRINTM", 4, 5, AccruedAtMaturity);
    }

    /// <summary>The bond price for a known yield (§18.17.7.257).</summary>
    internal static double BondPrice(in CouponSchedule schedule, double rate, double yield, double redemption)
    {
        int frequency = schedule.Frequency;
        double coupon = 100 * rate / frequency;
        double periodYield = 1 + (yield / frequency);
        double toNext = schedule.DaysToNext / schedule.PeriodLength;
        int count = schedule.Count;

        if (count == 1)
        {
            return ((redemption + coupon) / (1 + (toNext * yield / frequency))) -
                (coupon * schedule.DaysAccrued / schedule.PeriodLength);
        }

        double price = redemption / Math.Pow(periodYield, count - 1 + toNext);
        for (int k = 1; k <= count; k++)
            price += coupon / Math.Pow(periodYield, k - 1 + toNext);
        return price - (coupon * schedule.DaysAccrued / schedule.PeriodLength);
    }

    private static FormulaValue Price(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadBond(context, args, out CouponSchedule schedule, out double rate, out double second,
                out double redemption, out CellValue error))
        {
            return FormulaValue.Scalar(error);
        }

        if (rate < 0 || second < 0 || redemption <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(BondPrice(schedule, rate, second, redemption));
    }

    private static FormulaValue Yield(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadBond(context, args, out CouponSchedule schedule, out double rate, out double price,
                out double redemption, out CellValue error))
        {
            return FormulaValue.Scalar(error);
        }

        if (rate < 0 || price <= 0 || redemption <= 0)
            return FormulaValue.Error(ExcelError.Number);

        // The price falls as the yield grows, so the root is bracketed and bisection cannot
        // miss it. The lower bound stops short of where 1+y/f crosses zero.
        double low = (-0.9999) * schedule.Frequency;
        double high = 100;
        if (BondPrice(schedule, rate, low, redemption) < price || BondPrice(schedule, rate, high, redemption) > price)
            return FormulaValue.Error(ExcelError.Number);

        for (int i = 0; i < 200; i++)
        {
            double middle = (low + high) / 2;
            if (BondPrice(schedule, rate, middle, redemption) > price)
                low = middle;
            else
                high = middle;
        }

        return FormulaValue.Number((low + high) / 2);
    }

    /// <summary>Macaulay's duration in years, or its modified form (§18.17.7.101, .220).</summary>
    private static FormulaValue Duration(EvalContext context, ReadOnlySpan<FormulaValue> args, bool modified)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !Args.TryNumber(context, args[2], out double rate, out error) ||
            !Args.TryNumber(context, args[3], out double yield, out error) ||
            !Args.TryInteger(context, args[4], out int frequency, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 5 && !args[5].IsBlank && !Args.TryInteger(context, args[5], out basis, out error))
            return FormulaValue.Scalar(error);

        if (rate < 0 || yield < 0 ||
            !CouponSchedule.TryCreate(settlement.Date, maturity.Date, frequency, basis, out CouponSchedule schedule))
        {
            return FormulaValue.Error(ExcelError.Number);
        }

        double coupon = 100 * rate / frequency;
        double periodYield = 1 + (yield / frequency);
        double toNext = schedule.DaysToNext / schedule.PeriodLength;

        double weighted = 0;
        double total = 0;
        for (int k = 1; k <= schedule.Count; k++)
        {
            double time = k - 1 + toNext;
            double flow = coupon / Math.Pow(periodYield, time);
            if (k == schedule.Count)
                flow += 100 / Math.Pow(periodYield, time);

            weighted += time * flow;
            total += flow;
        }

        if (total == 0)
            return FormulaValue.Error(ExcelError.Number);

        double duration = weighted / total / frequency;
        return FormulaValue.Number(modified ? duration / periodYield : duration);
    }

    private static FormulaValue Disc(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Discounted(context, args, static (price, redemption, fraction) =>
            (1 - (price / redemption)) / fraction);

    private static FormulaValue PriceDisc(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Discounted(context, args, static (discount, redemption, fraction) =>
            redemption * (1 - (discount * fraction)));

    private static FormulaValue YieldDisc(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Discounted(context, args, static (price, redemption, fraction) =>
            ((redemption / price) - 1) / fraction);

    private static FormulaValue IntRate(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Discounted(context, args, static (investment, redemption, fraction) =>
            ((redemption / investment) - 1) / fraction);

    private static FormulaValue Received(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Discounted(context, args, static (investment, discount, fraction) =>
            investment / (1 - (discount * fraction)));

    /// <summary>
    /// The shared shape of the discounted family: settlement, maturity, two numbers and a
    /// basis, combined over the year fraction between the dates.
    /// </summary>
    private static FormulaValue Discounted(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Func<double, double, double, double> combine)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !Args.TryNumber(context, args[2], out double first, out error) ||
            !Args.TryNumber(context, args[3], out double second, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryInteger(context, args[4], out basis, out error))
            return FormulaValue.Scalar(error);

        if (settlement >= maturity || first <= 0 || second <= 0 || basis is < 0 or > 4)
            return FormulaValue.Error(ExcelError.Number);

        double fraction = DayCountFunctions.YearFraction(settlement.Date, maturity.Date, basis);
        double result = combine(first, second, fraction);
        return double.IsNaN(result) || double.IsInfinity(result) || result < 0
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(result);
    }

    /// <summary>Price of a security paying all its interest at maturity (§18.17.7.259).</summary>
    private static FormulaValue PriceMat(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadMaturityInterest(context, args, out DateTime settlement, out DateTime maturity,
                out DateTime issue, out double rate, out double second, out int basis, out CellValue error))
        {
            return FormulaValue.Scalar(error);
        }

        double issueToMaturity = DayCountFunctions.YearFraction(issue, maturity, basis);
        double issueToSettlement = DayCountFunctions.YearFraction(issue, settlement, basis);
        double settlementToMaturity = DayCountFunctions.YearFraction(settlement, maturity, basis);

        double price = ((100 + (issueToMaturity * rate * 100)) / (1 + (settlementToMaturity * second))) -
            (issueToSettlement * rate * 100);
        return FormulaValue.Number(price);
    }

    /// <summary>Yield of the same security, the closed inversion of the price (§18.17.7.353).</summary>
    private static FormulaValue YieldMat(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryReadMaturityInterest(context, args, out DateTime settlement, out DateTime maturity,
                out DateTime issue, out double rate, out double price, out int basis, out CellValue error))
        {
            return FormulaValue.Scalar(error);
        }

        if (price <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double issueToMaturity = DayCountFunctions.YearFraction(issue, maturity, basis);
        double issueToSettlement = DayCountFunctions.YearFraction(issue, settlement, basis);
        double settlementToMaturity = DayCountFunctions.YearFraction(settlement, maturity, basis);

        double invested = (price / 100) + (issueToSettlement * rate);
        double yield = ((1 + (issueToMaturity * rate)) - invested) / invested / settlementToMaturity;
        return FormulaValue.Number(yield);
    }

    /// <summary>settlement, maturity, issue, rate, price-or-yield, [basis] — shared by the pair.</summary>
    private static bool TryReadMaturityInterest(
        EvalContext context, ReadOnlySpan<FormulaValue> args, out DateTime settlement, out DateTime maturity,
        out DateTime issue, out double rate, out double second, out int basis, out CellValue error)
    {
        settlement = maturity = issue = default;
        rate = second = 0;
        basis = 0;
        if (!DateTimeFunctions.TryDate(context, args[0], out settlement, out error) ||
            !DateTimeFunctions.TryDate(context, args[1], out maturity, out error) ||
            !DateTimeFunctions.TryDate(context, args[2], out issue, out error) ||
            !Args.TryNumber(context, args[3], out rate, out error) ||
            !Args.TryNumber(context, args[4], out second, out error))
        {
            return false;
        }

        if (args.Length > 5 && !args[5].IsBlank && !Args.TryInteger(context, args[5], out basis, out error))
            return false;

        settlement = settlement.Date;
        maturity = maturity.Date;
        issue = issue.Date;
        if (settlement >= maturity || issue > settlement || rate < 0 || basis is < 0 or > 4)
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        return true;
    }

    private enum BillKind
    {
        Price,
        Yield,
        Equivalent,
    }

    /// <summary>The Treasury bill trio, all on actual days over a 360-day year.</summary>
    private static FormulaValue Bill(EvalContext context, ReadOnlySpan<FormulaValue> args, BillKind kind)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !Args.TryNumber(context, args[2], out double argument, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int days = (maturity.Date - settlement.Date).Days;
        if (days <= 0 || days > 366 || argument <= 0)
            return FormulaValue.Error(ExcelError.Number);

        switch (kind)
        {
            case BillKind.Price:
                double price = 100 * (1 - (argument * days / 360));
                return price <= 0 ? FormulaValue.Error(ExcelError.Number) : FormulaValue.Number(price);

            case BillKind.Yield:
                return FormulaValue.Number((100 - argument) / argument * 360 / days);

            default:
                double denominator = 360 - (argument * days);
                return denominator <= 0
                    ? FormulaValue.Error(ExcelError.Number)
                    : FormulaValue.Number(365 * argument / denominator);
        }
    }

    /// <summary>
    /// <c>ACCRINT</c> — interest accrued on a periodic security, walked over the quasi-coupon
    /// periods that the stretch from issue to settlement crosses (§18.17.7.2).
    /// </summary>
    private static FormulaValue AccruedPeriodic(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime issue, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime firstInterest, out error) ||
            !DateTimeFunctions.TryDate(context, args[2], out DateTime settlement, out error) ||
            !Args.TryNumber(context, args[3], out double rate, out error) ||
            !Args.TryNumber(context, args[4], out double par, out error) ||
            !Args.TryInteger(context, args[5], out int frequency, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 6 && !args[6].IsBlank && !Args.TryInteger(context, args[6], out basis, out error))
            return FormulaValue.Scalar(error);

        bool fromIssue = true;
        if (args.Length > 7 && !args[7].IsBlank && !Args.TryBoolean(context, args[7], out fromIssue, out error))
            return FormulaValue.Scalar(error);

        issue = issue.Date;
        firstInterest = firstInterest.Date;
        settlement = settlement.Date;
        if (rate <= 0 || par <= 0 || frequency is not (1 or 2 or 4) || basis is < 0 or > 4 || issue >= settlement)
            return FormulaValue.Error(ExcelError.Number);

        // The calculation method only matters once settlement has passed the first interest
        // date; before that both readings accrue from issue.
        DateTime from = fromIssue || settlement <= firstInterest ? issue : firstInterest;
        if (from >= settlement)
            return FormulaValue.Number(0);

        return FormulaValue.Number(
            par * rate / frequency * QuasiPeriodFractions(from, settlement, firstInterest, frequency, basis));
    }

    /// <summary>
    /// Sums the fraction each quasi-coupon period contributes between two dates, on a grid
    /// anchored at the first interest date.
    /// </summary>
    private static double QuasiPeriodFractions(
        DateTime from, DateTime to, DateTime anchor, int frequency, int basis)
    {
        int step = 12 / frequency;

        // Walk the anchor back to the period holding the start; a start past the anchor is
        // reached by the forward march below instead.
        DateTime periodEnd = anchor;
        while (CouponSchedule.AddMonthsPinned(periodEnd, -step) > from)
            periodEnd = CouponSchedule.AddMonthsPinned(periodEnd, -step);

        double total = 0;
        DateTime periodStart = CouponSchedule.AddMonthsPinned(periodEnd, -step);
        while (periodStart < to)
        {
            DateTime accrueFrom = periodStart < from ? from : periodStart;
            DateTime accrueTo = periodEnd < to ? periodEnd : to;
            if (accrueTo > accrueFrom)
            {
                double days = CountDays(accrueFrom, accrueTo, basis);
                double length = basis switch
                {
                    1 => (periodEnd - periodStart).Days,
                    3 => 365.0 / frequency,
                    _ => 360.0 / frequency,
                };
                total += days / length;
            }

            periodStart = periodEnd;
            periodEnd = CouponSchedule.AddMonthsPinned(periodEnd, step);
        }

        return total;
    }

    private static double CountDays(DateTime from, DateTime to, int basis) => basis switch
    {
        0 => DayCountFunctions.Days360Nasd(from, to),
        4 => DayCountFunctions.Days360European(from, to),
        _ => (to - from).Days,
    };

    /// <summary><c>ACCRINTM</c> — the same accrual in one stretch, issue to settlement (§18.17.7.3).</summary>
    private static FormulaValue AccruedAtMaturity(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime issue, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime settlement, out error) ||
            !Args.TryNumber(context, args[2], out double rate, out error) ||
            !Args.TryNumber(context, args[3], out double par, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryInteger(context, args[4], out basis, out error))
            return FormulaValue.Scalar(error);

        if (rate <= 0 || par <= 0 || basis is < 0 or > 4 || issue.Date >= settlement.Date)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(par * rate * DayCountFunctions.YearFraction(issue.Date, settlement.Date, basis));
    }

    /// <summary>settlement, maturity, rate, yield-or-price, redemption, frequency, [basis].</summary>
    private static bool TryReadBond(
        EvalContext context, ReadOnlySpan<FormulaValue> args, out CouponSchedule schedule,
        out double rate, out double second, out double redemption, out CellValue error)
    {
        schedule = default;
        rate = second = redemption = 0;
        if (!DateTimeFunctions.TryDate(context, args[0], out DateTime settlement, out error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime maturity, out error) ||
            !Args.TryNumber(context, args[2], out rate, out error) ||
            !Args.TryNumber(context, args[3], out second, out error) ||
            !Args.TryNumber(context, args[4], out redemption, out error) ||
            !Args.TryInteger(context, args[5], out int frequency, out error))
        {
            return false;
        }

        int basis = 0;
        if (args.Length > 6 && !args[6].IsBlank && !Args.TryInteger(context, args[6], out basis, out error))
            return false;

        if (!CouponSchedule.TryCreate(settlement.Date, maturity.Date, frequency, basis, out schedule))
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        return true;
    }
}
