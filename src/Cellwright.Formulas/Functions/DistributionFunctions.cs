namespace Cellwright.Formulas.Functions;

/// <summary>
/// The continuous distributions: normal, log-normal, Student's t, chi-squared, F, beta, gamma,
/// exponential and Weibull.
/// </summary>
/// <remarks>
/// Each is a thin reading of <see cref="SpecialFunctions"/>. Excel renamed this whole catalog
/// in 2010 — <c>NORMDIST</c> became <c>NORM.DIST</c> and so on — but kept the old names working,
/// and files thirty years old are full of them, so both spellings are registered against the
/// same implementation. Where an argument leaves a distribution undefined the answer is
/// <c>#NUM!</c>, as it is in Excel.
/// </remarks>
internal static class DistributionFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("NORM.DIST", 4, 4, Normal);
        builder.Add("NORMDIST", 4, 4, Normal);
        builder.Add("NORM.INV", 3, 3, NormalInverse);
        builder.Add("NORMINV", 3, 3, NormalInverse);
        builder.Add("NORM.S.DIST", 1, 2, StandardNormal);
        builder.Add("NORMSDIST", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.NormalBelow));
        builder.Add("NORM.S.INV", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.NormalQuantile));
        builder.Add("NORMSINV", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.NormalQuantile));
        builder.Add("PHI", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.NormalDensity));
        builder.Add("GAUSS", 1, 1, (c, a) => Args.Map(c, a[0], static z => SpecialFunctions.NormalBelow(z) - 0.5));
        builder.Add("STANDARDIZE", 3, 3, Standardize);

        builder.Add("LOGNORM.DIST", 4, 4, LogNormal);
        builder.Add("LOGNORMDIST", 3, 3, (c, a) => LogNormalBelow(c, a));
        builder.Add("LOGNORM.INV", 3, 3, LogNormalInverse);
        builder.Add("LOGINV", 3, 3, LogNormalInverse);

        builder.Add("EXPON.DIST", 3, 3, Exponential);
        builder.Add("EXPONDIST", 3, 3, Exponential);
        builder.Add("WEIBULL.DIST", 4, 4, Weibull);
        builder.Add("WEIBULL", 4, 4, Weibull);

        builder.Add("ERF", 1, 2, ErrorFunction);
        builder.Add("ERF.PRECISE", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.Erf));
        builder.Add("ERFC", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.Erfc));
        builder.Add("ERFC.PRECISE", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.Erfc));

        builder.Add("GAMMA", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.Gamma));
        builder.Add("GAMMALN", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.LogGamma));
        builder.Add("GAMMALN.PRECISE", 1, 1, (c, a) => Args.Map(c, a[0], SpecialFunctions.LogGamma));
        builder.Add("GAMMA.DIST", 4, 4, GammaDistribution);
        builder.Add("GAMMADIST", 4, 4, GammaDistribution);
        builder.Add("GAMMA.INV", 3, 3, GammaInverse);
        builder.Add("GAMMAINV", 3, 3, GammaInverse);

        builder.Add("BETA.DIST", 4, 6, BetaDistribution);
        builder.Add("BETADIST", 3, 5, BetaCumulative);
        builder.Add("BETA.INV", 3, 5, BetaInverse);
        builder.Add("BETAINV", 3, 5, BetaInverse);

        StudentFunctions.Register(builder);
        ChiSquareFunctions.Register(builder);
        DiscreteDistributionFunctions.Register(builder);
    }

    /// <summary>
    /// <c>ERF(lower, [upper])</c>: the error function, or the span of it between two points
    /// when a second is given.
    /// </summary>
    private static FormulaValue ErrorFunction(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double lower, out CellValue error))
            return FormulaValue.Scalar(error);
        if (args.Length < 2 || args[1].IsBlank)
            return Finite(SpecialFunctions.Erf(lower));

        if (!Args.TryNumber(context, args[1], out double upper, out error))
            return FormulaValue.Scalar(error);

        return Finite(SpecialFunctions.Erf(upper) - SpecialFunctions.Erf(lower));
    }

    private static FormulaValue Normal(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryNumber(context, args[2], out double deviation, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (deviation <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double z = (x - mean) / deviation;
        return Finite(cumulative
            ? SpecialFunctions.NormalBelow(z)
            : SpecialFunctions.NormalDensity(z) / deviation);
    }

    private static FormulaValue NormalInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryNumber(context, args[2], out double deviation, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (deviation <= 0 || probability is <= 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);

        return Finite(mean + (deviation * SpecialFunctions.NormalQuantile(probability)));
    }

    private static FormulaValue StandardNormal(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double z, out CellValue error))
            return FormulaValue.Scalar(error);

        bool cumulative = true;
        if (args.Length > 1 && !Args.TryBoolean(context, args[1], out cumulative, out error))
            return FormulaValue.Scalar(error);

        return Finite(cumulative ? SpecialFunctions.NormalBelow(z) : SpecialFunctions.NormalDensity(z));
    }

    private static FormulaValue Standardize(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryNumber(context, args[2], out double deviation, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return deviation <= 0
            ? FormulaValue.Error(ExcelError.Number)
            : Finite((x - mean) / deviation);
    }

    private static FormulaValue LogNormal(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryNumber(context, args[2], out double deviation, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (deviation <= 0 || x <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double z = (Math.Log(x) - mean) / deviation;
        return Finite(cumulative
            ? SpecialFunctions.NormalBelow(z)
            : SpecialFunctions.NormalDensity(z) / (x * deviation));
    }

    /// <summary>The 1990s spelling, which only ever had the cumulative form.</summary>
    private static FormulaValue LogNormalBelow(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<FormulaValue> extended = [args[0], args[1], args[2], FormulaValue.Boolean(true)];
        return LogNormal(context, extended);
    }

    private static FormulaValue LogNormalInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryNumber(context, args[2], out double deviation, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (deviation <= 0 || probability is <= 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);

        return Finite(Math.Exp(mean + (deviation * SpecialFunctions.NormalQuantile(probability))));
    }

    private static FormulaValue Exponential(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double rate, out error) ||
            !Args.TryBoolean(context, args[2], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (x < 0 || rate <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return Finite(cumulative ? 1 - Math.Exp(-rate * x) : rate * Math.Exp(-rate * x));
    }

    private static FormulaValue Weibull(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double shape, out error) ||
            !Args.TryNumber(context, args[2], out double scale, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (x < 0 || shape <= 0 || scale <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double ratio = Math.Pow(x / scale, shape);
        return Finite(cumulative
            ? 1 - Math.Exp(-ratio)
            : shape / Math.Pow(scale, shape) * Math.Pow(x, shape - 1) * Math.Exp(-ratio));
    }

    private static FormulaValue GammaDistribution(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double shape, out error) ||
            !Args.TryNumber(context, args[2], out double scale, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (x < 0 || shape <= 0 || scale <= 0)
            return FormulaValue.Error(ExcelError.Number);

        double z = x / scale;
        return Finite(cumulative
            ? SpecialFunctions.LowerGamma(shape, z)
            : Math.Exp(((shape - 1) * Math.Log(z)) - z - SpecialFunctions.LogGamma(shape)) / scale);
    }

    private static FormulaValue GammaInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double shape, out error) ||
            !Args.TryNumber(context, args[2], out double scale, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (shape <= 0 || scale <= 0 || probability is < 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);
        if (probability == 0)
            return FormulaValue.Number(0);

        double z = SpecialFunctions.Invert(t => SpecialFunctions.LowerGamma(shape, t), probability, 0, shape + 1);
        return Finite(z * scale);
    }

    /// <summary>
    /// <c>BETA.DIST(x, alpha, beta, cumulative, [a], [b])</c>. The optional bounds rescale the
    /// distribution from the unit interval it is defined on to the one the caller cares about.
    /// </summary>
    private static FormulaValue BetaDistribution(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double alpha, out error) ||
            !Args.TryNumber(context, args[2], out double beta, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (!TryBounds(context, args, 4, out double low, out double high, out error))
            return FormulaValue.Scalar(error);

        if (alpha <= 0 || beta <= 0 || high <= low || x < low || x > high)
            return FormulaValue.Error(ExcelError.Number);

        double scaled = (x - low) / (high - low);
        if (cumulative)
            return Finite(SpecialFunctions.IncompleteBeta(alpha, beta, scaled));

        double density = Math.Exp(
            ((alpha - 1) * Math.Log(scaled)) + ((beta - 1) * Math.Log(1 - scaled)) +
            SpecialFunctions.LogGamma(alpha + beta) - SpecialFunctions.LogGamma(alpha) - SpecialFunctions.LogGamma(beta));

        return Finite(density / (high - low));
    }

    /// <summary>The older spelling, which is the cumulative form and nothing else.</summary>
    private static FormulaValue BetaCumulative(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<FormulaValue> extended =
        [
            args[0], args[1], args[2], FormulaValue.Boolean(true),
            Args.At(args, 3), Args.At(args, 4),
        ];

        return BetaDistribution(context, extended);
    }

    private static FormulaValue BetaInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double probability, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double alpha, out error) ||
            !Args.TryNumber(context, args[2], out double beta, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (!TryBounds(context, args, 3, out double low, out double high, out error))
            return FormulaValue.Scalar(error);

        if (alpha <= 0 || beta <= 0 || high <= low || probability is < 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        double scaled = SpecialFunctions.Invert(t => SpecialFunctions.IncompleteBeta(alpha, beta, t), probability, 0, 1);
        if (probability == 0 || probability == 1)
            scaled = probability;

        return Finite(low + (scaled * (high - low)));
    }

    /// <summary>Reads the optional pair of bounds a beta distribution may be stretched onto.</summary>
    private static bool TryBounds(
        EvalContext context, ReadOnlySpan<FormulaValue> args, int index,
        out double low, out double high, out CellValue error)
    {
        low = 0;
        high = 1;
        error = CellValue.Blank;
        if (args.Length > index && !args[index].IsBlank && !Args.TryNumber(context, args[index], out low, out error))
            return false;

        return args.Length <= index + 1 || args[index + 1].IsBlank ||
            Args.TryNumber(context, args[index + 1], out high, out error);
    }

    /// <summary>A result that ran off the end of the number line is <c>#NUM!</c>, not infinity.</summary>
    internal static FormulaValue Finite(double value) =>
        double.IsNaN(value) || double.IsInfinity(value)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(value);
}
