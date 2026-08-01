namespace Cellwright.Formulas.Functions;

/// <summary>
/// The financial functions that need no coupon calendar: rate conversions, the cumulative
/// halves of an annuity, the dollar-fraction pair, a variable declining balance and the two
/// French depreciations.
/// </summary>
internal static class AnnualRateFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("EFFECT", 2, 2, Effect);
        builder.Add("NOMINAL", 2, 2, Nominal);
        builder.Add("CUMIPMT", 6, 6, (c, a) => Cumulative(c, a, interest: true));
        builder.Add("CUMPRINC", 6, 6, (c, a) => Cumulative(c, a, interest: false));
        builder.Add("ISPMT", 4, 4, IsPmt);
        builder.Add("FVSCHEDULE", 2, 2, FvSchedule, wantsRawReferences: true);
        builder.Add("DOLLARDE", 2, 2, (c, a) => DollarFraction(c, a, toDecimal: true));
        builder.Add("DOLLARFR", 2, 2, (c, a) => DollarFraction(c, a, toDecimal: false));
        builder.Add("VDB", 5, 7, VariableDecliningBalance);
        builder.Add("AMORLINC", 6, 7, (c, a) => Amortization(c, a, degressive: false));
        builder.Add("AMORDEGRC", 6, 7, (c, a) => Amortization(c, a, degressive: true));
    }

    private static FormulaValue Effect(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double nominal, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double periodsRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double periods = Math.Truncate(periodsRaw);
        if (nominal <= 0 || periods < 1)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(Math.Pow(1 + (nominal / periods), periods) - 1);
    }

    private static FormulaValue Nominal(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double effective, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double periodsRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double periods = Math.Truncate(periodsRaw);
        if (effective <= 0 || periods < 1)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number((Math.Pow(1 + effective, 1 / periods) - 1) * periods);
    }

    /// <summary>
    /// <c>CUMIPMT</c> and <c>CUMPRINC</c> — the interest or principal halves of the annuity
    /// payment, added up over a span of periods. Walked balance by balance: the closed forms
    /// exist, but the walk is exact and the spans are human-sized.
    /// </summary>
    private static FormulaValue Cumulative(EvalContext context, ReadOnlySpan<FormulaValue> args, bool interest)
    {
        if (!Args.TryNumber(context, args[0], out double rate, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double nper, out error) ||
            !Args.TryNumber(context, args[2], out double pv, out error) ||
            !Args.TryInteger(context, args[3], out int start, out error) ||
            !Args.TryInteger(context, args[4], out int end, out error) ||
            !Args.TryInteger(context, args[5], out int type, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || type is not (0 or 1))
            return FormulaValue.Error(ExcelError.Number);

        // The standard annuity payment, both halves of which this walk splits.
        double factor = Math.Pow(1 + rate, nper);
        double payment = -pv * rate * factor / ((factor - 1) * (1 + (rate * type)));

        double balance = pv;
        double total = 0;
        for (int period = 1; period <= end; period++)
        {
            double interestPart = period == 1 && type == 1 ? 0 : -balance * rate;
            double principalPart = payment - interestPart;
            if (period >= start)
                total += interest ? interestPart : principalPart;
            balance += principalPart;
        }

        return FormulaValue.Number(total);
    }

    /// <summary>The even-principal loan's interest for one period (§18.17.7.181).</summary>
    private static FormulaValue IsPmt(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double rate, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double period, out error) ||
            !Args.TryNumber(context, args[2], out double nper, out error) ||
            !Args.TryNumber(context, args[3], out double pv, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (nper == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        return FormulaValue.Number(-pv * rate * (1 - (period / nper)));
    }

    private static FormulaValue FvSchedule(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double principal, out CellValue error))
            return FormulaValue.Scalar(error);

        foreach (CellValue value in new ArgumentValues(args[1], context))
        {
            if (value.IsError)
                return FormulaValue.Scalar(value);
            if (value.IsBlank)
                continue;
            if (value.IsText)
                return FormulaValue.Error(ExcelError.Value);

            principal *= 1 + value.ToSerialNumber(context.DateSystem);
        }

        return FormulaValue.Number(principal);
    }

    /// <summary>
    /// <c>DOLLARDE</c> reads <c>1.02</c> as one and two thirty-seconds when the base is 32;
    /// <c>DOLLARFR</c> writes it back. The fraction digits scale by the base's width.
    /// </summary>
    private static FormulaValue DollarFraction(EvalContext context, ReadOnlySpan<FormulaValue> args, bool toDecimal)
    {
        if (!Args.TryNumber(context, args[0], out double amount, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double baseRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double fraction = Math.Truncate(baseRaw);
        if (fraction < 0)
            return FormulaValue.Error(ExcelError.Number);
        if (fraction == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double digits = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
        double whole = Math.Truncate(amount);
        double part = amount - whole;
        double converted = toDecimal
            ? whole + (part * digits / fraction)
            : whole + (part * fraction / digits);
        return FormulaValue.Number(converted);
    }

    /// <summary>
    /// <c>VDB</c> — declining balance over any span, including fractional edges, switching to
    /// straight line the moment straight line depreciates faster (unless told not to).
    /// </summary>
    private static FormulaValue VariableDecliningBalance(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double salvage, out error) ||
            !Args.TryNumber(context, args[2], out double life, out error) ||
            !Args.TryNumber(context, args[3], out double start, out error) ||
            !Args.TryNumber(context, args[4], out double end, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double factor = 2;
        if (args.Length > 5 && !args[5].IsBlank && !Args.TryNumber(context, args[5], out factor, out error))
            return FormulaValue.Scalar(error);

        bool noSwitch = false;
        if (args.Length > 6 && !args[6].IsBlank && !Args.TryBoolean(context, args[6], out noSwitch, out error))
            return FormulaValue.Scalar(error);

        if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor < 0)
            return FormulaValue.Error(ExcelError.Number);

        double book = cost;
        double total = 0;
        int lastPeriod = (int)Math.Ceiling(end);
        for (int period = 0; period < lastPeriod; period++)
        {
            double declining = Math.Min(book * factor / life, book - salvage);
            double straight = life - period > 0 ? (book - salvage) / (life - period) : 0;
            double depreciation = !noSwitch && straight > declining ? straight : declining;
            depreciation = Math.Max(Math.Min(depreciation, book - salvage), 0);

            double covered = Math.Min(period + 1, end) - Math.Max(period, start);
            if (covered > 0)
                total += depreciation * covered;
            book -= depreciation;
        }

        return FormulaValue.Number(total);
    }

    /// <summary>
    /// The French accounting depreciations. <c>AMORLINC</c> spreads the rate linearly with a
    /// prorated first period; <c>AMORDEGRC</c> multiplies the rate by the life's coefficient,
    /// rounds every figure, and folds the tail into a final half.
    /// </summary>
    private static FormulaValue Amortization(EvalContext context, ReadOnlySpan<FormulaValue> args, bool degressive)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !DateTimeFunctions.TryDate(context, args[1], out DateTime purchased, out error) ||
            !DateTimeFunctions.TryDate(context, args[2], out DateTime firstPeriod, out error) ||
            !Args.TryNumber(context, args[3], out double salvage, out error) ||
            !Args.TryInteger(context, args[4], out int period, out error) ||
            !Args.TryNumber(context, args[5], out double rate, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int basis = 0;
        if (args.Length > 6 && !args[6].IsBlank && !Args.TryInteger(context, args[6], out basis, out error))
            return FormulaValue.Scalar(error);

        if (cost < 0 || salvage < 0 || rate <= 0 || period < 0 || basis is < 0 or > 4 || basis == 2)
            return FormulaValue.Error(ExcelError.Number);

        double firstFraction = DayCountFunctions.YearFraction(purchased, firstPeriod, basis);
        if (!degressive)
        {
            double yearly = cost * rate;
            double first = yearly * firstFraction;
            double depreciable = cost - salvage;
            if (period == 0)
                return FormulaValue.Number(Math.Min(first, depreciable));

            double accumulated = Math.Min(first, depreciable);
            for (int i = 1; i < period; i++)
                accumulated += Math.Min(yearly, Math.Max(depreciable - accumulated, 0));

            return FormulaValue.Number(Math.Min(yearly, Math.Max(depreciable - accumulated, 0)));
        }

        // The degressive coefficient grows with the asset's life, and every amount is rounded
        // to whole currency the way the French rules state.
        double usefulLife = 1 / rate;
        if (usefulLife is >= 0 and < 3)
            rate *= 1;
        else if (usefulLife < 5)
            rate *= 1.5;
        else if (usefulLife <= 6)
            rate *= 2;
        else
            rate *= 2.5;

        double book = cost;
        double depreciation = Math.Round(firstFraction * rate * cost, MidpointRounding.AwayFromZero);
        if (period == 0)
            return FormulaValue.Number(depreciation);

        book -= depreciation;
        double rest = book - salvage;
        for (int i = 1; i <= period; i++)
        {
            depreciation = Math.Round(rate * book, MidpointRounding.AwayFromZero);
            rest -= depreciation;

            // The book value runs out: the asked-for period gets half the remainder, periods
            // beyond it get nothing.
            if (rest < 0)
                return FormulaValue.Number(i == period ? Math.Round(book * 0.5, MidpointRounding.AwayFromZero) : 0);

            book -= depreciation;
        }

        return FormulaValue.Number(depreciation);
    }
}
