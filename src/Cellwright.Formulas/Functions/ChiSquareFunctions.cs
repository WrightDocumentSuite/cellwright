namespace Cellwright.Formulas.Functions;

/// <summary>
/// The two distributions of a test statistic: chi-squared, and the F of a ratio of variances.
/// </summary>
/// <remarks>
/// Both are nearly always read from the right, because a test asks how unlikely a statistic at
/// least this large would be. Excel's older names give the right tail and its newer ones the
/// left, which is a trap worth naming: <c>CHIDIST</c> and <c>CHISQ.DIST</c> are not the same
/// function.
/// </remarks>
internal static class ChiSquareFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("CHISQ.DIST", 3, 3, ChiSquare);
        builder.Add("CHISQ.DIST.RT", 2, 2, ChiSquareRight);
        builder.Add("CHIDIST", 2, 2, ChiSquareRight);
        builder.Add("CHISQ.INV", 2, 2, (c, a) => ChiSquareInverse(c, a, fromRight: false));
        builder.Add("CHISQ.INV.RT", 2, 2, (c, a) => ChiSquareInverse(c, a, fromRight: true));
        builder.Add("CHIINV", 2, 2, (c, a) => ChiSquareInverse(c, a, fromRight: true));

        builder.Add("F.DIST", 4, 4, FisherDistribution);
        builder.Add("F.DIST.RT", 3, 3, FisherRight);
        builder.Add("FDIST", 3, 3, FisherRight);
        builder.Add("F.INV", 3, 3, (c, a) => FisherInverse(c, a, fromRight: false));
        builder.Add("F.INV.RT", 3, 3, (c, a) => FisherInverse(c, a, fromRight: true));
        builder.Add("FINV", 3, 3, (c, a) => FisherInverse(c, a, fromRight: true));
    }

    /// <summary>The share of a chi-squared distribution lying below <paramref name="x"/>.</summary>
    private static double Below(double x, double freedom) => SpecialFunctions.LowerGamma(freedom / 2, x / 2);

    private static FormulaValue ChiSquare(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error) ||
            !Args.TryBoolean(context, args[2], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (x < 0 || freedom < 1)
            return FormulaValue.Error(ExcelError.Number);
        if (cumulative)
            return DistributionFunctions.Finite(Below(x, freedom));

        double half = freedom / 2;
        double density = Math.Exp(
            ((half - 1) * Math.Log(x)) - (x / 2) - (half * Math.Log(2)) - SpecialFunctions.LogGamma(half));

        return DistributionFunctions.Finite(density);
    }

    private static FormulaValue ChiSquareRight(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (x < 0 || freedom < 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(SpecialFunctions.UpperGamma(freedom / 2, x / 2));
    }

    private static FormulaValue ChiSquareInverse(EvalContext context, ReadOnlySpan<FormulaValue> args, bool fromRight)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (freedom < 1 || probability is < 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        double below = fromRight ? 1 - probability : probability;
        if (below <= 0)
            return FormulaValue.Number(0);
        if (below >= 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(
            SpecialFunctions.Invert(x => Below(x, freedom), below, 0, freedom + 2));
    }

    /// <summary>The share of an F distribution lying below <paramref name="x"/>.</summary>
    private static double FisherBelow(double x, double first, double second) =>
        SpecialFunctions.IncompleteBeta(first / 2, second / 2, first * x / ((first * x) + second));

    private static FormulaValue FisherDistribution(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryRatio(context, args, out double x, out double first, out double second, out CellValue error))
            return FormulaValue.Scalar(error);
        if (!Args.TryBoolean(context, args[3], out bool cumulative, out error))
            return FormulaValue.Scalar(error);

        if (cumulative)
            return DistributionFunctions.Finite(FisherBelow(x, first, second));
        if (x == 0)
            return FormulaValue.Number(first > 2 ? 0 : double.PositiveInfinity);

        double logDensity =
            (0.5 * ((first * Math.Log(first * x)) + (second * Math.Log(second)) -
                ((first + second) * Math.Log((first * x) + second)))) -
            Math.Log(x) -
            (SpecialFunctions.LogGamma(first / 2) + SpecialFunctions.LogGamma(second / 2) -
                SpecialFunctions.LogGamma((first + second) / 2));

        return DistributionFunctions.Finite(Math.Exp(logDensity));
    }

    private static FormulaValue FisherRight(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryRatio(context, args, out double x, out double first, out double second, out CellValue error))
            return FormulaValue.Scalar(error);

        return DistributionFunctions.Finite(1 - FisherBelow(x, first, second));
    }

    private static FormulaValue FisherInverse(EvalContext context, ReadOnlySpan<FormulaValue> args, bool fromRight)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double first, out error) ||
            !Args.TryNumber(context, args[2], out double second, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (first < 1 || second < 1 || probability is < 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        double below = fromRight ? 1 - probability : probability;
        if (below <= 0)
            return FormulaValue.Number(0);
        if (below >= 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(
            SpecialFunctions.Invert(x => FisherBelow(x, first, second), below, 0, 4));
    }

    /// <summary>Reads the point and the two degrees of freedom an F distribution is written with.</summary>
    private static bool TryRatio(
        EvalContext context, ReadOnlySpan<FormulaValue> args,
        out double x, out double first, out double second, out CellValue error)
    {
        first = 0;
        second = 0;
        if (!Args.TryNumber(context, args[0], out x, out error) ||
            !Args.TryNumber(context, args[1], out first, out error) ||
            !Args.TryNumber(context, args[2], out second, out error))
        {
            return false;
        }

        if (x < 0 || first < 1 || second < 1)
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        return true;
    }
}
