namespace Cellwright.Formulas.Functions;

/// <summary>
/// Counting and descriptive statistics.
/// </summary>
internal static class StatisticalFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("COUNT", 1, FunctionDefinition.Unlimited, Count, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("COUNTA", 1, FunctionDefinition.Unlimited, CountA, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("COUNTBLANK", 1, 1, CountBlank, wantsRawReferences: true);
        builder.Add("AVERAGE", 1, FunctionDefinition.Unlimited, Average, wantsRawReferences: true);
        builder.Add("MIN", 1, FunctionDefinition.Unlimited, (c, a) => Extreme(c, a, min: true), wantsRawReferences: true);
        builder.Add("MAX", 1, FunctionDefinition.Unlimited, (c, a) => Extreme(c, a, min: false), wantsRawReferences: true);
        builder.Add("MEDIAN", 1, FunctionDefinition.Unlimited, Median, wantsRawReferences: true);
        builder.Add("MODE", 1, FunctionDefinition.Unlimited, Mode, wantsRawReferences: true);
        builder.Add("LARGE", 2, 2, (c, a) => Nth(c, a, largest: true), wantsRawReferences: true);
        builder.Add("SMALL", 2, 2, (c, a) => Nth(c, a, largest: false), wantsRawReferences: true);
        builder.Add("RANK", 2, 3, Rank, wantsRawReferences: true);
        builder.Add("STDEV", 1, FunctionDefinition.Unlimited, (c, a) => Spread(c, a, sample: true, root: true), wantsRawReferences: true);
        builder.Add("STDEVP", 1, FunctionDefinition.Unlimited, (c, a) => Spread(c, a, sample: false, root: true), wantsRawReferences: true);
        builder.Add("VAR", 1, FunctionDefinition.Unlimited, (c, a) => Spread(c, a, sample: true, root: false), wantsRawReferences: true);
        builder.Add("VARP", 1, FunctionDefinition.Unlimited, (c, a) => Spread(c, a, sample: false, root: false), wantsRawReferences: true);
        builder.Add("PERCENTILE", 2, 2, Percentile, wantsRawReferences: true);
        builder.Add("QUARTILE", 2, 2, Quartile, wantsRawReferences: true);
        builder.Add("PERCENTILE.INC", 2, 2, Percentile, wantsRawReferences: true);
        builder.Add("QUARTILE.INC", 2, 2, Quartile, wantsRawReferences: true);
        builder.Add("PERCENTILE.EXC", 2, 2, PercentileExclusive, wantsRawReferences: true);
        builder.Add("QUARTILE.EXC", 2, 2, QuartileExclusive, wantsRawReferences: true);
    }

    private static FormulaValue Count(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        int count = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsNumber || value.IsDateTime || value.IsTimeSpan)
                    count++;
                else if (!fromRange && !value.IsError && !value.IsBlank && Coercion.ToNumber(value, context.DateSystem).IsNumber)
                    count++;
            }
        }

        return FormulaValue.Number(count);
    }

    private static FormulaValue CountA(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        int count = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!value.IsBlank)
                    count++;
            }
        }

        return FormulaValue.Number(count);
    }

    private static FormulaValue CountBlank(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        int count = 0;
        foreach (CellValue value in new ArgumentValues(args[0], context))
        {
            if (value.IsBlank || (value.IsText && value.GetText().Length == 0))
                count++;
        }

        return FormulaValue.Number(count);
    }

    private static FormulaValue Average(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        Aggregates.Accumulator stats = Aggregates.Collect(context, args);
        if (stats.Failed)
            return FormulaValue.Scalar(stats.Error);
        return stats.Count == 0
            ? FormulaValue.Error(ExcelError.DivisionByZero)
            : FormulaValue.Number(stats.Sum / stats.Count);
    }

    private static FormulaValue Extreme(EvalContext context, ReadOnlySpan<FormulaValue> args, bool min)
    {
        Aggregates.Accumulator stats = Aggregates.Collect(context, args);
        if (stats.Failed)
            return FormulaValue.Scalar(stats.Error);
        if (stats.Count == 0)
            return FormulaValue.Number(0);
        return FormulaValue.Number(min ? stats.Min : stats.Max);
    }

    private static FormulaValue Median(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> numbers = Collect(context, args, out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);
        if (numbers.Count == 0)
            return FormulaValue.Error(ExcelError.Number);

        numbers.Sort();
        int middle = numbers.Count / 2;
        double median = numbers.Count % 2 == 1
            ? numbers[middle]
            : (numbers[middle - 1] + numbers[middle]) / 2;
        return FormulaValue.Number(median);
    }

    private static FormulaValue Mode(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        List<double> numbers = Collect(context, args, out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);

        var counts = new Dictionary<double, int>();
        double best = 0;
        int bestCount = 1;
        foreach (double number in numbers)
        {
            int seen = counts.TryGetValue(number, out int existing) ? existing + 1 : 1;
            counts[number] = seen;
            if (seen > bestCount)
            {
                bestCount = seen;
                best = number;
            }
        }

        return bestCount > 1 ? FormulaValue.Number(best) : FormulaValue.Error(ExcelError.NotAvailable);
    }

    private static FormulaValue Nth(EvalContext context, ReadOnlySpan<FormulaValue> args, bool largest)
    {
        List<double> numbers = Collect(context, args[..1], out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);
        if (!Args.TryInteger(context, args[1], out int k, out error))
            return FormulaValue.Scalar(error);
        if (k < 1 || k > numbers.Count)
            return FormulaValue.Error(ExcelError.Number);

        numbers.Sort();
        return FormulaValue.Number(largest ? numbers[^k] : numbers[k - 1]);
    }

    private static FormulaValue Rank(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double target, out CellValue error))
            return FormulaValue.Scalar(error);

        List<double> numbers = Collect(context, args[1..2], out error);
        if (error.IsError)
            return FormulaValue.Scalar(error);

        bool ascending = false;
        if (args.Length > 2)
        {
            if (!Args.TryNumber(context, args[2], out double order, out error))
                return FormulaValue.Scalar(error);
            ascending = order != 0;
        }

        int rank = 1;
        bool found = false;
        foreach (double number in numbers)
        {
            if (number == target)
                found = true;
            else if (ascending ? number < target : number > target)
                rank++;
        }

        return found ? FormulaValue.Number(rank) : FormulaValue.Error(ExcelError.NotAvailable);
    }

    private static FormulaValue Spread(EvalContext context, ReadOnlySpan<FormulaValue> args, bool sample, bool root)
    {
        List<double> numbers = Collect(context, args, out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);

        int divisor = sample ? numbers.Count - 1 : numbers.Count;
        if (divisor <= 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double mean = numbers.Sum() / numbers.Count;
        double sumOfSquares = 0;
        foreach (double number in numbers)
            sumOfSquares += (number - mean) * (number - mean);

        double variance = sumOfSquares / divisor;
        return FormulaValue.Number(root ? Math.Sqrt(variance) : variance);
    }

    private static FormulaValue Percentile(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[1], out double fraction, out CellValue error))
            return FormulaValue.Scalar(error);
        return PercentileCore(context, args[..1], fraction);
    }

    private static FormulaValue Quartile(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[1], out int quart, out CellValue error))
            return FormulaValue.Scalar(error);
        if (quart is < 0 or > 4)
            return FormulaValue.Error(ExcelError.Number);
        return PercentileCore(context, args[..1], quart / 4.0);
    }

    private static FormulaValue PercentileExclusive(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[1], out double fraction, out CellValue error))
            return FormulaValue.Scalar(error);
        return PercentileExclusiveCore(context, args[..1], fraction);
    }

    private static FormulaValue QuartileExclusive(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[1], out int quart, out CellValue error))
            return FormulaValue.Scalar(error);

        // The exclusive form has no minimum or maximum, so quarts 0 and 4 have no answer.
        if (quart is < 1 or > 3)
            return FormulaValue.Error(ExcelError.Number);
        return PercentileExclusiveCore(context, args[..1], quart / 4.0);
    }

    private static FormulaValue PercentileCore(EvalContext context, ReadOnlySpan<FormulaValue> args, double fraction)
    {
        if (fraction is < 0 or > 1)
            return FormulaValue.Error(ExcelError.Number);

        List<double> numbers = Collect(context, args, out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);

        // The standard leaves an empty array undefined; Excel answers zero (MS-OI29500
        // 2.1.1014). The exclusive form below has no such rule and still refuses.
        if (numbers.Count == 0)
            return FormulaValue.Number(0);

        numbers.Sort();
        double position = fraction * (numbers.Count - 1);
        int lower = (int)Math.Floor(position);
        int upper = (int)Math.Ceiling(position);
        double weight = position - lower;
        return FormulaValue.Number(numbers[lower] + ((numbers[upper] - numbers[lower]) * weight));
    }

    /// <summary>
    /// The exclusive variants rank the sample at <c>k*(n+1)</c> instead of <c>k*(n-1)</c>, so
    /// the ends of the data are never reached: a fraction outside <c>1/(n+1) … n/(n+1)</c>
    /// has no interpolation and yields #NUM!.
    /// </summary>
    private static FormulaValue PercentileExclusiveCore(EvalContext context, ReadOnlySpan<FormulaValue> args, double fraction)
    {
        List<double> numbers = Collect(context, args, out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);
        if (numbers.Count == 0)
            return FormulaValue.Error(ExcelError.Number);

        double rank = fraction * (numbers.Count + 1);
        if (rank < 1 || rank > numbers.Count)
            return FormulaValue.Error(ExcelError.Number);

        numbers.Sort();
        int lower = (int)Math.Floor(rank);
        if (lower >= numbers.Count)
            return FormulaValue.Number(numbers[^1]);

        double weight = rank - lower;
        return FormulaValue.Number(numbers[lower - 1] + ((numbers[lower] - numbers[lower - 1]) * weight));
    }

    /// <summary>Gathers every numeric value of the arguments into a list for order statistics.</summary>
    private static List<double> Collect(EvalContext context, ReadOnlySpan<FormulaValue> args, out CellValue error)
    {
        var numbers = new List<double>();
        error = CellValue.Blank;
        foreach (ref readonly FormulaValue arg in args)
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (Aggregates.TryNumeric(value, fromRange, context, out double number, out CellValue failure))
                    numbers.Add(number);
                else if (failure.IsError)
                {
                    error = failure;
                    return numbers;
                }
            }
        }

        return numbers;
    }
}
