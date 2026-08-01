namespace Cellwright.Formulas.Functions;

/// <summary>
/// Student's t distribution and the confidence intervals built on it.
/// </summary>
/// <remarks>
/// The t family is unusual in having three cumulative forms — left tail, right tail and both
/// tails at once — because a t test is nearly always asking about a tail rather than about the
/// distribution itself. All three come off the same incomplete beta.
/// </remarks>
internal static class StudentFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("T.DIST", 3, 3, Student);
        builder.Add("T.DIST.RT", 2, 2, (c, a) => Tail(c, a, doubled: false));
        builder.Add("T.DIST.2T", 2, 2, (c, a) => Tail(c, a, doubled: true));
        builder.Add("TDIST", 3, 3, LegacyStudent);
        builder.Add("T.INV", 2, 2, StudentInverse);
        builder.Add("T.INV.2T", 2, 2, TwoTailedInverse);
        builder.Add("TINV", 2, 2, TwoTailedInverse);
        builder.Add("CONFIDENCE", 3, 3, (c, a) => Confidence(c, a, student: false));
        builder.Add("CONFIDENCE.NORM", 3, 3, (c, a) => Confidence(c, a, student: false));
        builder.Add("CONFIDENCE.T", 3, 3, (c, a) => Confidence(c, a, student: true));
    }

    /// <summary>The share of a t distribution lying below <paramref name="t"/>.</summary>
    public static double Below(double t, double freedom)
    {
        double half = SpecialFunctions.IncompleteBeta(freedom / 2, 0.5, freedom / (freedom + (t * t))) / 2;
        return t >= 0 ? 1 - half : half;
    }

    private static double Density(double t, double freedom) => Math.Exp(
        SpecialFunctions.LogGamma((freedom + 1) / 2) - SpecialFunctions.LogGamma(freedom / 2) -
        (0.5 * Math.Log(freedom * Math.PI)) -
        ((freedom + 1) / 2 * Math.Log(1 + (t * t / freedom))));

    private static FormulaValue Student(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double t, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error) ||
            !Args.TryBoolean(context, args[2], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (freedom < 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(cumulative ? Below(t, freedom) : Density(t, freedom));
    }

    /// <summary>The right tail, or both tails together; either way the point must not be negative.</summary>
    private static FormulaValue Tail(EvalContext context, ReadOnlySpan<FormulaValue> args, bool doubled)
    {
        if (!Args.TryNumber(context, args[0], out double t, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (freedom < 1 || (doubled && t < 0))
            return FormulaValue.Error(ExcelError.Number);

        double right = 1 - Below(t, freedom);
        return DistributionFunctions.Finite(doubled ? 2 * (1 - Below(Math.Abs(t), freedom)) : right);
    }

    /// <summary>
    /// <c>TDIST(x, df, tails)</c>, which predates the split into named forms and takes the
    /// number of tails as an argument. It has never accepted a negative point.
    /// </summary>
    private static FormulaValue LegacyStudent(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double t, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error) ||
            !Args.TryInteger(context, args[2], out int tails, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (t < 0 || freedom < 1 || tails is not (1 or 2))
            return FormulaValue.Error(ExcelError.Number);

        double right = 1 - Below(t, freedom);
        return DistributionFunctions.Finite(tails == 1 ? right : 2 * right);
    }

    private static FormulaValue StudentInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (freedom < 1 || probability is <= 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(Quantile(probability, freedom));
    }

    /// <summary>
    /// <c>T.INV.2T(p, df)</c>: the point beyond which, in either direction, a share
    /// <paramref name="args"/> of the distribution lies. Always positive by construction.
    /// </summary>
    private static FormulaValue TwoTailedInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double freedom, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (freedom < 1 || probability is <= 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(Quantile(1 - (probability / 2), freedom));
    }

    /// <summary>
    /// The point below which a given share lies. The distribution is symmetric, so only the
    /// upper half is searched and the lower half is its mirror.
    /// </summary>
    private static double Quantile(double probability, double freedom)
    {
        if (probability == 0.5)
            return 0;
        if (probability < 0.5)
            return -Quantile(1 - probability, freedom);

        return SpecialFunctions.Invert(t => Below(t, freedom), probability, 0, 4);
    }

    /// <summary>
    /// Half the width of a confidence interval about a mean. The normal form assumes the spread
    /// of the population is known; the t form allows for its having been estimated from the
    /// sample, which is why it needs one fewer degree of freedom than there are observations.
    /// </summary>
    private static FormulaValue Confidence(EvalContext context, ReadOnlySpan<FormulaValue> args, bool student)
    {
        if (!Args.TryNumber(context, args[0], out double alpha, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double deviation, out error) ||
            !Args.TryNumber(context, args[2], out double size, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (alpha is <= 0 or >= 1 || deviation <= 0 || size < 1 || (student && size < 2))
            return FormulaValue.Error(ExcelError.Number);

        double spread = student
            ? Quantile(1 - (alpha / 2), Math.Floor(size) - 1)
            : SpecialFunctions.NormalQuantile(1 - (alpha / 2));

        return DistributionFunctions.Finite(spread * deviation / Math.Sqrt(Math.Floor(size)));
    }
}
