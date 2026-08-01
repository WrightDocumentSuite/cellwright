namespace Cellwright.Formulas.Functions;

/// <summary>
/// The four ways of writing an asset down: evenly, by the sum of the years, and the two
/// declining balances.
/// </summary>
/// <remarks>
/// A declining balance is defined period by period, since each period's charge is taken from
/// what the last one left. Both are worked out in closed form here rather than by walking the
/// years, because the balance falls geometrically until it meets the salvage value and stops —
/// and a life stated in days, as <c>DDB(cost, salvage, 10*365, 1)</c> states it, would be a
/// long walk for one number.
/// </remarks>
internal static class DepreciationFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SLN", 3, 3, StraightLine);
        builder.Add("SYD", 4, 4, SumOfYears);
        builder.Add("DDB", 4, 5, DecliningBalance);
        builder.Add("DB", 4, 5, FixedDecliningBalance);
    }

    private static FormulaValue StraightLine(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double salvage, out error) ||
            !Args.TryNumber(context, args[2], out double life, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return life == 0
            ? FormulaValue.Error(ExcelError.DivisionByZero)
            : Finite((cost - salvage) / life);
    }

    /// <summary>
    /// <c>SYD</c> spreads the writedown over the years in the proportion each year's number
    /// bears to the sum of them all, so the first year takes the largest share.
    /// </summary>
    private static FormulaValue SumOfYears(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double salvage, out error) ||
            !Args.TryNumber(context, args[2], out double life, out error) ||
            !Args.TryNumber(context, args[3], out double period, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (life <= 0 || period < 1 || period > life)
            return FormulaValue.Error(ExcelError.Number);

        return Finite((cost - salvage) * (life - period + 1) * 2 / (life * (life + 1)));
    }

    /// <summary>
    /// <c>DDB(cost, salvage, life, period, [factor])</c>: a constant share of whatever is left,
    /// stopping the moment the total written down reaches the salvage value.
    /// </summary>
    private static FormulaValue DecliningBalance(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double salvage, out error) ||
            !Args.TryNumber(context, args[2], out double life, out error) ||
            !Args.TryNumber(context, args[3], out double period, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double factor = 2;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryNumber(context, args[4], out factor, out error))
            return FormulaValue.Scalar(error);

        if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || period > life || factor <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double rate = Math.Min(factor / life, 1);

        // A part period takes that part of the charge the whole one would have carried.
        double whole = Math.Ceiling(period);
        double share = period - whole + 1;
        double remaining = cost * Math.Pow(1 - rate, whole - 1);
        double charge = Math.Min(remaining * rate, remaining - salvage);
        return Finite(Math.Max(charge, 0) * share);
    }

    /// <summary>
    /// <c>DB(cost, salvage, life, period, [month])</c>: the rate is fixed once, from the ground
    /// the asset has to lose over its life, and rounded to three places — which is why this and
    /// <c>DDB</c> disagree even when their rates would otherwise match.
    /// </summary>
    private static FormulaValue FixedDecliningBalance(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double cost, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double salvage, out error) ||
            !Args.TryNumber(context, args[2], out double life, out error) ||
            !Args.TryNumber(context, args[3], out double period, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double month = 12;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryNumber(context, args[4], out month, out error))
            return FormulaValue.Scalar(error);

        month = Math.Floor(month);
        period = Math.Floor(period);
        if (cost < 0 || salvage < 0 || life < 1 || period < 1 || period > life + 1 || month is < 1 or > 12)
            return FormulaValue.Error(ExcelError.Number);
        if (cost == 0)
            return FormulaValue.Number(0);

        double rate = Math.Round(1 - Math.Pow(salvage / cost, 1 / life), 3, MidpointRounding.AwayFromZero);

        // The first period is short by however much of the year had already gone, and the last
        // one picks up the remainder of that same year.
        double first = cost * rate * month / 12;
        if (period == 1)
            return Finite(first);

        double left = cost - first;
        double charge = left * Math.Pow(1 - rate, period - 2) * rate;
        return Finite(period > life ? charge * (12 - month) / 12 : charge);
    }

    private static FormulaValue Finite(double value) =>
        double.IsNaN(value) || double.IsInfinity(value)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(value);
}
