namespace Cellwright.Formulas.Functions;

/// <summary>
/// Statistical functions beyond the base catalog. <c>FREQUENCY</c> lives here; the
/// correlation and hypothesis-test families share the file's pairing helpers.
/// </summary>
internal static class ExtendedStatisticalFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("FREQUENCY", 2, 2, Frequency, wantsRawReferences: true);
        builder.Add("AVEDEV", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.AverageDeviation), wantsRawReferences: true);
        builder.Add("DEVSQ", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.SquaredDeviation), wantsRawReferences: true);
        builder.Add("GEOMEAN", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.GeometricMean), wantsRawReferences: true);
        builder.Add("HARMEAN", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.HarmonicMean), wantsRawReferences: true);
        builder.Add("KURT", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.Kurtosis), wantsRawReferences: true);
        builder.Add("SKEW", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.Skewness), wantsRawReferences: true);
        builder.Add("SKEW.P", 1, FunctionDefinition.Unlimited, (c, a) => Moments(c, a, MomentKind.SkewnessPopulation), wantsRawReferences: true);
        builder.Add("TRIMMEAN", 2, 2, TrimMean, wantsRawReferences: true);
        builder.Add("FISHER", 1, 1, (c, a) => Args.Map(c, a[0], static x => Math.Abs(x) < 1 ? 0.5 * Math.Log((1 + x) / (1 - x)) : double.NaN));
        builder.Add("FISHERINV", 1, 1, (c, a) => Args.Map(c, a[0], static y => Math.Tanh(y)));
        builder.Add("PERCENTRANK", 2, 3, (c, a) => PercentRank(c, a, exclusive: false), wantsRawReferences: true);
        builder.Alias("PERCENTRANK.INC", "PERCENTRANK");
        builder.Add("PERCENTRANK.EXC", 2, 3, (c, a) => PercentRank(c, a, exclusive: true), wantsRawReferences: true);
        builder.Add("PROB", 3, 4, Probability, wantsRawReferences: true);
        builder.Add("MODE.MULT", 1, FunctionDefinition.Unlimited, ModeMultiple, wantsRawReferences: true);
        builder.Add("RANK.AVG", 2, 3, RankAverage, wantsRawReferences: true);
        builder.Add("STDEVA", 1, FunctionDefinition.Unlimited, (c, a) => CountingEverything(c, a, sample: true, root: true), wantsRawReferences: true);
        builder.Add("STDEVPA", 1, FunctionDefinition.Unlimited, (c, a) => CountingEverything(c, a, sample: false, root: true), wantsRawReferences: true);
        builder.Add("VARA", 1, FunctionDefinition.Unlimited, (c, a) => CountingEverything(c, a, sample: true, root: false), wantsRawReferences: true);
        builder.Add("VARPA", 1, FunctionDefinition.Unlimited, (c, a) => CountingEverything(c, a, sample: false, root: false), wantsRawReferences: true);
    }

    private enum MomentKind
    {
        AverageDeviation,
        SquaredDeviation,
        GeometricMean,
        HarmonicMean,
        Kurtosis,
        Skewness,
        SkewnessPopulation,
    }

    /// <summary>The single-sample statistics that reduce to moments around the mean.</summary>
    private static FormulaValue Moments(EvalContext context, ReadOnlySpan<FormulaValue> args, MomentKind kind)
    {
        List<double> values = [];
        foreach (ref readonly FormulaValue arg in args)
        {
            if (CollectNumbers(context, arg, values) is { IsError: true } error)
                return FormulaValue.Scalar(error);
        }

        int n = values.Count;
        if (n == 0)
            return FormulaValue.Error(kind is MomentKind.GeometricMean or MomentKind.HarmonicMean
                ? ExcelError.Number
                : ExcelError.DivisionByZero);

        if (kind == MomentKind.GeometricMean)
        {
            double logSum = 0;
            foreach (double value in values)
            {
                if (value <= 0)
                    return FormulaValue.Error(ExcelError.Number);
                logSum += Math.Log(value);
            }

            return FormulaValue.Number(Math.Exp(logSum / n));
        }

        if (kind == MomentKind.HarmonicMean)
        {
            double reciprocals = 0;
            foreach (double value in values)
            {
                if (value <= 0)
                    return FormulaValue.Error(ExcelError.Number);
                reciprocals += 1 / value;
            }

            return FormulaValue.Number(n / reciprocals);
        }

        double mean = 0;
        foreach (double value in values)
            mean += value;
        mean /= n;

        double second = 0;
        double third = 0;
        double fourth = 0;
        double absolute = 0;
        foreach (double value in values)
        {
            double d = value - mean;
            absolute += Math.Abs(d);
            second += d * d;
            third += d * d * d;
            fourth += d * d * d * d;
        }

        switch (kind)
        {
            case MomentKind.AverageDeviation:
                return FormulaValue.Number(absolute / n);
            case MomentKind.SquaredDeviation:
                return FormulaValue.Number(second);
            case MomentKind.Skewness:
                {
                    if (n < 3 || second == 0)
                        return FormulaValue.Error(ExcelError.DivisionByZero);
                    double s = Math.Sqrt(second / (n - 1));
                    return FormulaValue.Number(n / ((n - 1.0) * (n - 2)) * (third / (s * s * s)));
                }

            case MomentKind.SkewnessPopulation:
                {
                    if (n < 3 || second == 0)
                        return FormulaValue.Error(ExcelError.DivisionByZero);
                    double sigma = Math.Sqrt(second / n);
                    return FormulaValue.Number(third / n / (sigma * sigma * sigma));
                }

            default:
                {
                    if (n < 4 || second == 0)
                        return FormulaValue.Error(ExcelError.DivisionByZero);
                    double s2 = second / (n - 1);
                    double a = (double)n * (n + 1) / ((n - 1.0) * (n - 2) * (n - 3));
                    double b = 3.0 * (n - 1) * (n - 1) / ((n - 2.0) * (n - 3));
                    return FormulaValue.Number((a * (fourth / (s2 * s2))) - b);
                }
        }
    }

    /// <summary>
    /// <c>TRIMMEAN(array, percent)</c> — the mean without the tails: <c>n·percent</c> points,
    /// rounded down to an even count, dropped half from each end.
    /// </summary>
    private static FormulaValue TrimMean(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> values = [];
        if (CollectNumbers(context, args[0], values) is { IsError: true } error)
            return FormulaValue.Scalar(error);
        if (!Args.TryNumber(context, args[1], out double percent, out CellValue bad))
            return FormulaValue.Scalar(bad);

        if (percent is < 0 or >= 1)
            return FormulaValue.Error(ExcelError.Number);
        if (values.Count == 0)
            return FormulaValue.Error(ExcelError.Number);

        values.Sort();
        int drop = (int)(values.Count * percent) / 2;
        double total = 0;
        int kept = 0;
        for (int i = drop; i < values.Count - drop; i++)
        {
            total += values[i];
            kept++;
        }

        return kept == 0 ? FormulaValue.Error(ExcelError.Number) : FormulaValue.Number(total / kept);
    }

    /// <summary>
    /// <c>PERCENTRANK</c> — where a value stands in its sample, interpolated between
    /// neighbours and cut (not rounded) to the asked number of digits.
    /// </summary>
    private static FormulaValue PercentRank(EvalContext context, ReadOnlySpan<FormulaValue> args, bool exclusive)
    {
        List<double> values = [];
        if (CollectNumbers(context, args[0], values) is { IsError: true } error)
            return FormulaValue.Scalar(error);
        if (!Args.TryNumber(context, args[1], out double x, out CellValue bad))
            return FormulaValue.Scalar(bad);

        int significance = 3;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryInteger(context, args[2], out significance, out bad))
            return FormulaValue.Scalar(bad);
        if (significance < 1)
            return FormulaValue.Error(ExcelError.Number);

        int n = values.Count;
        if (n == 0)
            return FormulaValue.Error(ExcelError.Number);

        values.Sort();
        if (x < values[0] || x > values[n - 1])
            return FormulaValue.Error(ExcelError.NotAvailable);

        int below = 0;
        while (below < n && values[below] < x)
            below++;

        double position;
        if (below < n && values[below] == x)
        {
            position = below;
        }
        else
        {
            double previous = values[below - 1];
            double next = values[below];
            position = below - 1 + ((x - previous) / (next - previous));
        }

        double rank = exclusive
            ? (position + 1) / (n + 1)
            : n == 1 ? 1 : position / (n - 1);

        double factor = Math.Pow(10, significance);
        return FormulaValue.Number(Math.Truncate(rank * factor) / factor);
    }

    /// <summary>
    /// <c>PROB(values, probabilities, lower, [upper])</c> — the probability mass between the
    /// two bounds. The masses must each sit in (0, 1] and add up to one.
    /// </summary>
    private static FormulaValue Probability(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> values = [];
        List<double> masses = [];
        if (CollectNumbers(context, args[0], values) is { IsError: true } error)
            return FormulaValue.Scalar(error);
        if (CollectNumbers(context, args[1], masses) is { IsError: true } massError)
            return FormulaValue.Scalar(massError);
        if (!Args.TryNumber(context, args[2], out double lower, out CellValue bad))
            return FormulaValue.Scalar(bad);

        double upper = lower;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryNumber(context, args[3], out upper, out bad))
            return FormulaValue.Scalar(bad);

        if (values.Count != masses.Count)
            return FormulaValue.Error(ExcelError.NotAvailable);

        double total = 0;
        double included = 0;
        for (int i = 0; i < values.Count; i++)
        {
            double mass = masses[i];
            if (mass <= 0 || mass > 1)
                return FormulaValue.Error(ExcelError.Number);
            total += mass;
            if (values[i] >= lower && values[i] <= upper)
                included += mass;
        }

        return Math.Abs(total - 1) > 1e-9
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(included);
    }

    /// <summary>
    /// <c>MODE.MULT</c> — every value that reaches the highest repeat count, in the order first
    /// met. Nothing repeating at all is <c>#N/A</c>, like the single-answer MODE.
    /// </summary>
    private static FormulaValue ModeMultiple(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> values = [];
        foreach (ref readonly FormulaValue arg in args)
        {
            if (CollectNumbers(context, arg, values) is { IsError: true } error)
                return FormulaValue.Scalar(error);
        }

        Dictionary<double, int> counts = [];
        List<double> order = [];
        int best = 1;
        foreach (double value in values)
        {
            System.Runtime.InteropServices.CollectionsMarshal.GetValueRefOrAddDefault(counts, value, out bool seen)++;
            if (!seen)
                order.Add(value);
            best = Math.Max(best, counts[value]);
        }

        if (best < 2)
            return FormulaValue.Error(ExcelError.NotAvailable);

        List<CellValue> modes = [];
        foreach (double value in order)
        {
            if (counts[value] == best)
                modes.Add(CellValue.Number(value));
        }

        return FormulaValue.Array(new ArrayValue(modes.Count, 1, [.. modes]));
    }

    /// <summary>
    /// <c>RANK.AVG</c> — like RANK, except a run of equal values shares the average of the
    /// ranks it spans.
    /// </summary>
    private static FormulaValue RankAverage(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error))
            return FormulaValue.Scalar(error);

        List<double> values = [];
        if (CollectNumbers(context, args[1], values) is { IsError: true } bad)
            return FormulaValue.Scalar(bad);

        bool ascending = false;
        if (args.Length > 2)
        {
            if (!Args.TryNumber(context, args[2], out double order, out CellValue orderError))
                return FormulaValue.Scalar(orderError);
            ascending = order != 0;
        }

        int better = 0;
        int ties = 0;
        foreach (double value in values)
        {
            if (value == x)
                ties++;
            else if (ascending ? value < x : value > x)
                better++;
        }

        if (ties == 0)
            return FormulaValue.Error(ExcelError.NotAvailable);

        return FormulaValue.Number(better + ((ties + 1) / 2.0));
    }

    /// <summary>
    /// The "A" spreads: text counts as zero and logicals as their number, the way AVERAGEA
    /// counts them, before the ordinary sample or population formulas apply.
    /// </summary>
    private static FormulaValue CountingEverything(
        EvalContext context, ReadOnlySpan<FormulaValue> args, bool sample, bool root)
    {
        List<double> values = [];
        foreach (ref readonly FormulaValue arg in args)
        {
            bool literalArray = arg.Kind == FormulaValueKind.Array;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsError)
                    return FormulaValue.Scalar(value);
                if (value.IsBlank || (literalArray && value.IsBoolean))
                    continue;

                values.Add(value.Kind switch
                {
                    CellValueKind.Number => value.GetNumber(),
                    CellValueKind.Boolean => value.GetBoolean() ? 1 : 0,
                    CellValueKind.DateTime or CellValueKind.TimeSpan => value.ToSerialNumber(context.DateSystem),
                    _ => 0,
                });
            }
        }

        int n = values.Count;
        if (sample ? n < 2 : n < 1)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double mean = 0;
        foreach (double value in values)
            mean += value;
        mean /= n;

        double spread = 0;
        foreach (double value in values)
            spread += (value - mean) * (value - mean);

        double variance = spread / (sample ? n - 1 : n);
        return FormulaValue.Number(root ? Math.Sqrt(variance) : variance);
    }

    /// <summary>
    /// <c>FREQUENCY(data, bins)</c> — one count per bin plus an overflow row. Each value lands
    /// in the bin with the smallest bound not below it, wherever that bin sits in the range:
    /// Excel sorts the bounds internally but reports the counts in the bins' own order.
    /// </summary>
    private static FormulaValue Frequency(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> data = [];
        if (CollectNumbers(context, args[0], data) is { IsError: true } error)
            return FormulaValue.Scalar(error);

        List<double> bins = [];
        if (CollectNumbers(context, args[1], bins) is { IsError: true } binError)
            return FormulaValue.Scalar(binError);

        var counts = new int[bins.Count + 1];
        if (bins.Count == 0)
        {
            counts[0] = data.Count;
        }
        else
        {
            // The bounds in sorted order, each remembering the row it reports at.
            var sorted = new (double Bound, int Index)[bins.Count];
            for (int i = 0; i < bins.Count; i++)
                sorted[i] = (bins[i], i);
            Array.Sort(sorted, static (a, b) => a.Bound.CompareTo(b.Bound));

            foreach (double value in data)
            {
                int slot = FindBin(sorted, value);
                counts[slot < 0 ? bins.Count : sorted[slot].Index]++;
            }
        }

        var values = new CellValue[counts.Length];
        for (int i = 0; i < counts.Length; i++)
            values[i] = CellValue.Number(counts[i]);
        return FormulaValue.Array(new ArrayValue(counts.Length, 1, values));
    }

    /// <summary>The sorted position of the smallest bound not below the value, or -1 past them all.</summary>
    private static int FindBin((double Bound, int Index)[] sorted, double value)
    {
        int low = 0;
        int high = sorted.Length - 1;
        int found = -1;
        while (low <= high)
        {
            int middle = (low + high) / 2;
            if (sorted[middle].Bound >= value)
            {
                found = middle;
                high = middle - 1;
            }
            else
            {
                low = middle + 1;
            }
        }

        return found;
    }

    /// <summary>
    /// Gathers the numbers of one argument, skipping what numeric aggregates skip. Returns the
    /// error that fails the whole function, or blank.
    /// </summary>
    internal static CellValue CollectNumbers(EvalContext context, in FormulaValue argument, List<double> into)
    {
        foreach (CellValue value in new ArgumentValues(argument, context))
        {
            if (!Aggregates.TryNumeric(value, argument.Kind != FormulaValueKind.Scalar, context, out double number, out CellValue error))
            {
                if (error.IsError)
                    return error;
                continue;
            }

            into.Add(number);
        }

        return CellValue.Blank;
    }
}
