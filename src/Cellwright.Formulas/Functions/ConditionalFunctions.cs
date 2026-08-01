namespace Cellwright.Formulas.Functions;

/// <summary>
/// Criteria-driven aggregates. A criterion is either a literal to match or a string carrying
/// a comparison, such as <c>"&gt;10"</c> or <c>"apple*"</c>.
/// </summary>
internal static class ConditionalFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SUMIF", 2, 3, (c, a) => SingleCriterion(c, a, Reduction.Sum), wantsRawReferences: true);
        builder.Add("AVERAGEIF", 2, 3, (c, a) => SingleCriterion(c, a, Reduction.Average), wantsRawReferences: true);
        builder.Add("COUNTIF", 2, 2, (c, a) => SingleCriterion(c, a, Reduction.Count), wantsRawReferences: true);
        builder.Add("SUMIFS", 3, FunctionDefinition.Unlimited, (c, a) => MultiCriteria(c, a, Reduction.Sum), wantsRawReferences: true);
        builder.Add("AVERAGEIFS", 3, FunctionDefinition.Unlimited, (c, a) => MultiCriteria(c, a, Reduction.Average), wantsRawReferences: true);
        builder.Add("MAXIFS", 3, FunctionDefinition.Unlimited, (c, a) => MultiCriteria(c, a, Reduction.Max), wantsRawReferences: true);
        builder.Add("MINIFS", 3, FunctionDefinition.Unlimited, (c, a) => MultiCriteria(c, a, Reduction.Min), wantsRawReferences: true);
        builder.Add("COUNTIFS", 2, FunctionDefinition.Unlimited, CountIfs, wantsRawReferences: true);
        builder.Add("SUMPRODUCT", 1, FunctionDefinition.Unlimited, SumProduct, wantsRawReferences: true);
    }

    private enum Reduction
    {
        Sum,
        Average,
        Count,
        Max,
        Min,
    }

    private static FormulaValue SingleCriterion(EvalContext context, ReadOnlySpan<FormulaValue> args, Reduction reduction)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue test))
            return FormulaValue.Error(ExcelError.Value);

        var criterion = Criterion.Parse(Operators.Dereference(args[1], context), context);
        ArrayValue? sum = null;
        if (reduction != Reduction.Count)
        {
            FormulaValue target = args.Length > 2 ? args[2] : args[0];
            if (!Operators.TryGetArray(target, context, out ArrayValue values))
                return FormulaValue.Error(ExcelError.Value);
            sum = values;
        }

        var accumulator = new Aggregates.Accumulator();
        int matches = 0;
        for (int i = 0; i < test.Count; i++)
        {
            if (!criterion.Matches(test.Values[i], context))
                continue;

            matches++;
            if (sum is null)
                continue;
            if (i < sum.Count && sum.Values[i].IsNumber)
                accumulator.Add(sum.Values[i].GetNumber());
        }

        return Reduce(reduction, accumulator, matches);
    }

    private static FormulaValue MultiCriteria(EvalContext context, ReadOnlySpan<FormulaValue> args, Reduction reduction)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue values))
            return FormulaValue.Error(ExcelError.Value);

        if (!TryBuildFilter(context, args[1..], values.Count, out bool[] included, out FormulaValue failure))
            return failure;

        var accumulator = new Aggregates.Accumulator();
        int matches = 0;
        for (int i = 0; i < values.Count; i++)
        {
            if (!included[i])
                continue;
            matches++;
            if (values.Values[i].IsNumber)
                accumulator.Add(values.Values[i].GetNumber());
        }

        return Reduce(reduction, accumulator, matches);
    }

    private static FormulaValue CountIfs(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue first))
            return FormulaValue.Error(ExcelError.Value);
        if (!TryBuildFilter(context, args, first.Count, out bool[] included, out FormulaValue failure))
            return failure;

        int count = 0;
        foreach (bool match in included)
        {
            if (match)
                count++;
        }

        return FormulaValue.Number(count);
    }

    /// <summary>
    /// Builds the mask of positions satisfying every (range, criterion) pair. Ranges must have
    /// the same shape, which Excel also requires.
    /// </summary>
    private static bool TryBuildFilter(
        EvalContext context, ReadOnlySpan<FormulaValue> pairs, int length, out bool[] included, out FormulaValue failure)
    {
        included = new bool[length];
        included.AsSpan().Fill(true);
        failure = FormulaValue.Blank;

        for (int pair = 0; pair + 1 < pairs.Length; pair += 2)
        {
            if (!Operators.TryGetArray(pairs[pair], context, out ArrayValue test) || test.Count != length)
            {
                failure = FormulaValue.Error(ExcelError.Value);
                return false;
            }

            var criterion = Criterion.Parse(Operators.Dereference(pairs[pair + 1], context), context);
            for (int i = 0; i < length; i++)
            {
                if (included[i] && !criterion.Matches(test.Values[i], context))
                    included[i] = false;
            }
        }

        return true;
    }

    private static FormulaValue Reduce(Reduction reduction, in Aggregates.Accumulator accumulator, int matches) => reduction switch
    {
        Reduction.Count => FormulaValue.Number(matches),
        Reduction.Average => accumulator.Count == 0
            ? FormulaValue.Error(ExcelError.DivisionByZero)
            : FormulaValue.Number(accumulator.Sum / accumulator.Count),
        Reduction.Max => FormulaValue.Number(accumulator.Count == 0 ? 0 : accumulator.Max),
        Reduction.Min => FormulaValue.Number(accumulator.Count == 0 ? 0 : accumulator.Min),
        _ => FormulaValue.Number(accumulator.Sum),
    };

    private static FormulaValue SumProduct(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue first))
            return FormulaValue.Error(ExcelError.Value);

        double total = 0;
        for (int i = 0; i < first.Count; i++)
        {
            double product = 1;
            for (int a = 0; a < args.Length; a++)
            {
                if (!Operators.TryGetArray(args[a], context, out ArrayValue operand) || operand.Count != first.Count)
                    return FormulaValue.Error(ExcelError.Value);

                CellValue value = operand.Values[i];
                if (value.IsError)
                    return FormulaValue.Scalar(value);

                product *= value.Kind switch
                {
                    CellValueKind.Number => value.GetNumber(),
                    CellValueKind.Boolean => value.GetBoolean() ? 1 : 0,
                    CellValueKind.DateTime or CellValueKind.TimeSpan => value.ToSerialNumber(context.DateSystem),
                    _ => 0,
                };
            }

            total += product;
        }

        return FormulaValue.Number(total);
    }
}
