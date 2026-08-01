namespace Cellwright.Formulas.Functions;

/// <summary>
/// The cash flow functions that work from dates rather than from even periods, and the modified
/// rate of return that allows for money costing one rate to borrow and earning another to lend.
/// </summary>
internal static class CashFlowFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("XNPV", 3, 3, ExactNetPresentValue, wantsRawReferences: true);
        builder.Add("XIRR", 2, 3, ExactInternalRate, wantsRawReferences: true);
        builder.Add("MIRR", 3, 3, ModifiedInternalRate, wantsRawReferences: true);
    }

    /// <summary>Days in the year these functions reckon by, whatever the calendar says.</summary>
    private const double Year = 365;

    /// <summary>
    /// <c>XNPV(rate, values, dates)</c>: the same discounting as <c>NPV</c>, except that each
    /// amount is discounted by the time actually elapsed rather than by its position in a list.
    /// </summary>
    private static FormulaValue ExactNetPresentValue(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double rate, out CellValue error))
            return FormulaValue.Scalar(error);
        if (!TrySchedule(context, args[1], args[2], out double[] amounts, out double[] days, out FormulaValue refused))
            return refused;

        // Discounting at a rate of zero or less is refused outright, where the standard places
        // no restriction on it at all (MS-OI29500 2.1.1071).
        if (rate <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return Finite(Worth(amounts, days, rate));
    }

    /// <summary>
    /// <c>XIRR(values, dates, [guess])</c>: the rate at which dated flows come to nothing. It
    /// needs money to have gone both ways, since a series that only ever gains has no such rate.
    /// </summary>
    private static FormulaValue ExactInternalRate(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TrySchedule(context, args[0], args[1], out double[] amounts, out double[] days, out FormulaValue refused))
            return refused;

        double guess = 0.1;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryNumber(context, args[2], out guess, out CellValue error))
            return FormulaValue.Scalar(error);

        if (!BothWays(amounts))
            return FormulaValue.Error(ExcelError.Number);

        return RateSolver.Solve(rate => rate <= -1 ? double.NaN : Worth(amounts, days, rate), guess);
    }

    /// <summary>
    /// <c>MIRR(values, finance_rate, reinvest_rate)</c>. The plain internal rate quietly assumes
    /// every sum taken out is put back at that same rate; this one asks what borrowing really
    /// costs and what the proceeds really earn, and needs no search to answer.
    /// </summary>
    private static FormulaValue ModifiedInternalRate(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryAmounts(context, args[0], out double[] amounts))
            return FormulaValue.Error(ExcelError.Value);
        if (!Args.TryNumber(context, args[1], out double borrowing, out CellValue error) ||
            !Args.TryNumber(context, args[2], out double lending, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int periods = amounts.Length;
        if (periods < 2 || !BothWays(amounts))
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double spent = 0;
        double earned = 0;
        for (int i = 0; i < periods; i++)
        {
            if (amounts[i] < 0)
                spent += amounts[i] / Math.Pow(1 + borrowing, i);
            else
                earned += amounts[i] * Math.Pow(1 + lending, periods - 1 - i);
        }

        if (spent == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        return Finite(Math.Pow(-earned / spent, 1.0 / (periods - 1)) - 1);
    }

    /// <summary>What the dated flows are worth today at a given rate.</summary>
    private static double Worth(double[] amounts, double[] days, double rate)
    {
        double total = 0;
        for (int i = 0; i < amounts.Length; i++)
            total += amounts[i] / Math.Pow(1 + rate, (days[i] - days[0]) / Year);

        return total;
    }

    /// <summary>Whether money went both out and in, without which there is no rate to find.</summary>
    private static bool BothWays(double[] amounts)
    {
        bool negative = false;
        bool positive = false;
        foreach (double amount in amounts)
        {
            negative |= amount < 0;
            positive |= amount > 0;
        }

        return negative && positive;
    }

    /// <summary>
    /// Reads the amounts and the days they fall on. The first date is the one everything is
    /// measured from, so nothing may fall before it, and the two ranges have to line up.
    /// </summary>
    private static bool TrySchedule(
        EvalContext context, in FormulaValue values, in FormulaValue schedule,
        out double[] amounts, out double[] days, out FormulaValue refused)
    {
        days = [];
        refused = FormulaValue.Error(ExcelError.Number);
        if (!TryAmounts(context, values, out amounts))
        {
            refused = FormulaValue.Error(ExcelError.Value);
            return false;
        }

        if (!Operators.TryGetArray(schedule, context, out ArrayValue dated) || dated.Count != amounts.Length)
            return false;

        days = new double[dated.Count];
        for (int i = 0; i < days.Length; i++)
        {
            if (!Aggregates.TryNumeric(dated.Values[i], fromRange: true, context, out days[i], out _))
                return false;

            days[i] = Math.Floor(days[i]);
            if (days[i] < days[0])
                return false;
        }

        return amounts.Length >= 2;
    }

    private static bool TryAmounts(EvalContext context, in FormulaValue values, out double[] amounts)
    {
        amounts = [];
        if (!Operators.TryGetArray(values, context, out ArrayValue array))
            return false;

        var collected = new List<double>(array.Count);
        foreach (CellValue value in array.Values)
        {
            if (Aggregates.TryNumeric(value, fromRange: true, context, out double amount, out _))
                collected.Add(amount);
        }

        amounts = [.. collected];
        return amounts.Length > 0;
    }

    private static FormulaValue Finite(double value) =>
        double.IsNaN(value) || double.IsInfinity(value)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(value);
}
