namespace Cellwright.Formulas.Functions;

/// <summary>
/// The distributions that count things: binomial, negative binomial, Poisson and hypergeometric.
/// </summary>
/// <remarks>
/// Each cumulative form could be a loop over its own probabilities, and for a handful of trials
/// that is what it amounts to. They are read off the incomplete beta and gamma instead, because
/// a binomial over a million trials is an ordinary thing to write in a spreadsheet and summing a
/// million terms to answer it is not.
/// </remarks>
internal static class DiscreteDistributionFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("BINOM.DIST", 4, 4, Binomial);
        builder.Add("BINOMDIST", 4, 4, Binomial);
        builder.Add("BINOM.INV", 3, 3, BinomialInverse);
        builder.Add("CRITBINOM", 3, 3, BinomialInverse);
        builder.Add("NEGBINOM.DIST", 4, 4, NegativeBinomial);
        builder.Add("NEGBINOMDIST", 3, 3, NegativeBinomialDensity);
        builder.Add("POISSON.DIST", 3, 3, Poisson);
        builder.Add("POISSON", 3, 3, Poisson);
        builder.Add("HYPGEOM.DIST", 4, 5, Hypergeometric);
        builder.Add("HYPGEOMDIST", 4, 4, Hypergeometric);
    }

    /// <summary>The logarithm of a binomial coefficient, which is a ratio of factorials.</summary>
    private static double LogChoose(double n, double k) =>
        SpecialFunctions.LogGamma(n + 1) - SpecialFunctions.LogGamma(k + 1) - SpecialFunctions.LogGamma(n - k + 1);

    private static FormulaValue Binomial(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double successes, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double trials, out error) ||
            !Args.TryNumber(context, args[2], out double chance, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        successes = Math.Floor(successes);
        trials = Math.Floor(trials);
        if (successes < 0 || successes > trials || chance is < 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        return DistributionFunctions.Finite(cumulative
            ? BinomialBelow(successes, trials, chance)
            : Mass(successes, trials, chance));
    }

    private static double Mass(double successes, double trials, double chance)
    {
        if (chance == 0)
            return successes == 0 ? 1 : 0;
        if (chance == 1)
            return successes == trials ? 1 : 0;

        return Math.Exp(
            LogChoose(trials, successes) +
            (successes * Math.Log(chance)) +
            ((trials - successes) * Math.Log(1 - chance)));
    }

    /// <summary>
    /// The chance of no more than <paramref name="successes"/> successes. Stated as an
    /// incomplete beta it costs the same whether the trials number ten or ten million.
    /// </summary>
    private static double BinomialBelow(double successes, double trials, double chance)
    {
        if (successes >= trials)
            return 1;
        if (chance == 0)
            return 1;
        if (chance == 1)
            return 0;

        return SpecialFunctions.IncompleteBeta(trials - successes, successes + 1, 1 - chance);
    }

    /// <summary>
    /// <c>BINOM.INV(trials, chance, alpha)</c>: the fewest successes whose chance of not being
    /// exceeded reaches <c>alpha</c>. The search walks up from the mean rather than from zero,
    /// since that is where the answer nearly always is.
    /// </summary>
    private static FormulaValue BinomialInverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double trials, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double chance, out error) ||
            !Args.TryNumber(context, args[2], out double alpha, out error))
        {
            return FormulaValue.Scalar(error);
        }

        trials = Math.Floor(trials);
        if (trials < 0 || chance is < 0 or > 1 || alpha is <= 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);

        long low = 0;
        long high = (long)trials;
        while (low < high)
        {
            long middle = low + ((high - low) / 2);
            if (BinomialBelow(middle, trials, chance) < alpha)
                low = middle + 1;
            else
                high = middle;
        }

        return FormulaValue.Number(low);
    }

    /// <summary>
    /// <c>NEGBINOM.DIST(failures, successes, chance, cumulative)</c>: how many failures come
    /// before a given number of successes.
    /// </summary>
    private static FormulaValue NegativeBinomial(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double failures, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double successes, out error) ||
            !Args.TryNumber(context, args[2], out double chance, out error) ||
            !Args.TryBoolean(context, args[3], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        failures = Math.Floor(failures);
        successes = Math.Floor(successes);
        if (failures < 0 || successes < 1 || chance is <= 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        if (cumulative)
            return DistributionFunctions.Finite(SpecialFunctions.IncompleteBeta(successes, failures + 1, chance));

        double density = Math.Exp(
            LogChoose(failures + successes - 1, successes - 1) +
            (successes * Math.Log(chance)) +
            (failures * Math.Log(1 - chance)));

        return DistributionFunctions.Finite(density);
    }

    /// <summary>The 1990s spelling, which has no cumulative form.</summary>
    private static FormulaValue NegativeBinomialDensity(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<FormulaValue> extended = [args[0], args[1], args[2], FormulaValue.Boolean(false)];
        return NegativeBinomial(context, extended);
    }

    private static FormulaValue Poisson(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double events, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double mean, out error) ||
            !Args.TryBoolean(context, args[2], out bool cumulative, out error))
        {
            return FormulaValue.Scalar(error);
        }

        events = Math.Floor(events);
        if (events < 0 || mean < 0)
            return FormulaValue.Error(ExcelError.Number);
        if (mean == 0)
            return FormulaValue.Number(events == 0 || cumulative ? 1 : 0);

        return DistributionFunctions.Finite(cumulative
            ? SpecialFunctions.UpperGamma(events + 1, mean)
            : Math.Exp((events * Math.Log(mean)) - mean - SpecialFunctions.LogGamma(events + 1)));
    }

    /// <summary>
    /// <c>HYPGEOM.DIST(drawn_successes, drawn, successes, population, [cumulative])</c>: drawing
    /// without replacement, where each draw changes the odds of the next.
    /// </summary>
    private static FormulaValue Hypergeometric(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Span<double> counts = stackalloc double[4];
        for (int i = 0; i < 4; i++)
        {
            if (!Args.TryNumber(context, args[i], out double value, out CellValue error))
                return FormulaValue.Scalar(error);
            counts[i] = Math.Floor(value);
        }

        bool cumulative = false;
        if (args.Length > 4 && !Args.TryBoolean(context, args[4], out cumulative, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        double drawnSuccesses = counts[0], drawn = counts[1], successes = counts[2], population = counts[3];
        if (drawnSuccesses < 0 || drawn <= 0 || successes <= 0 || population <= 0 ||
            drawn > population || successes > population ||
            drawnSuccesses > drawn || drawnSuccesses > successes ||
            drawnSuccesses < drawn - (population - successes))
        {
            return FormulaValue.Error(ExcelError.Number);
        }

        if (!cumulative)
            return DistributionFunctions.Finite(Draw(drawnSuccesses, drawn, successes, population));

        // The tail is short by construction — it cannot run past the smaller of the two counts —
        // so summing it outright is cheap and exact.
        double total = 0;
        for (double k = Math.Max(0, drawn - (population - successes)); k <= drawnSuccesses; k++)
            total += Draw(k, drawn, successes, population);

        return DistributionFunctions.Finite(total);
    }

    private static double Draw(double drawnSuccesses, double drawn, double successes, double population) =>
        Math.Exp(
            LogChoose(successes, drawnSuccesses) +
            LogChoose(population - successes, drawn - drawnSuccesses) -
            LogChoose(population, drawn));
}
