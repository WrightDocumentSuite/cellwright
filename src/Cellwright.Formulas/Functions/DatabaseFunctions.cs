namespace Cellwright.Formulas.Functions;

/// <summary>
/// The database functions: one aggregate each, over the records a criteria table admits. They
/// predate <c>SUMIFS</c> and its relatives and say the same things a longer way round, but
/// files written any time in the last thirty years are full of them.
/// </summary>
internal static class DatabaseFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        Add(builder, "DSUM", Reduction.Sum);
        Add(builder, "DAVERAGE", Reduction.Average);
        Add(builder, "DCOUNT", Reduction.Count);
        Add(builder, "DCOUNTA", Reduction.CountA);
        Add(builder, "DMAX", Reduction.Max);
        Add(builder, "DMIN", Reduction.Min);
        Add(builder, "DPRODUCT", Reduction.Product);
        Add(builder, "DGET", Reduction.Get);
        Add(builder, "DSTDEV", Reduction.Stdev);
        Add(builder, "DSTDEVP", Reduction.StdevP);
        Add(builder, "DVAR", Reduction.Var);
        Add(builder, "DVARP", Reduction.VarP);
    }

    private static void Add(FunctionRegistryBuilder builder, string name, Reduction reduction) =>
        builder.Add(name, 3, 3, (c, a) => Apply(c, a, reduction), wantsRawReferences: true);

    private enum Reduction
    {
        Sum,
        Average,
        Count,
        CountA,
        Max,
        Min,
        Product,
        Get,
        Stdev,
        StdevP,
        Var,
        VarP,
    }

    private static FormulaValue Apply(EvalContext context, ReadOnlySpan<FormulaValue> args, Reduction reduction)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue database) ||
            !Operators.TryGetArray(args[2], context, out ArrayValue criteria))
        {
            return FormulaValue.Error(ExcelError.Value);
        }

        if (!DatabaseQuery.TryBuild(database, criteria, context, out DatabaseQuery query))
            return FormulaValue.Error(ExcelError.Value);

        // Counting records needs no field at all, and that is the one case where the middle
        // argument may be left out: DCOUNTA(base,,conditions) asks how many records got through.
        // A field of zero says the same thing, there being no zeroth column to have meant.
        bool counting = reduction is Reduction.Count or Reduction.CountA;
        CellValue field = Operators.Dereference(args[1], context);
        if (counting && (args[1].IsBlank || (field.IsNumber && field.GetNumber() == 0)))
            return FormulaValue.Number(Admitted(query, context));

        if (!DatabaseQuery.TryResolveField(database, field, context, out int column))
            return FormulaValue.Error(ExcelError.Value);

        return Collect(query, column, reduction, context);
    }

    private static int Admitted(in DatabaseQuery query, EvalContext context)
    {
        int count = 0;
        for (int record = 0; record < query.Records; record++)
        {
            if (query.Admits(record, context))
                count++;
        }

        return count;
    }

    private static FormulaValue Collect(
        in DatabaseQuery query, int column, Reduction reduction, EvalContext context)
    {
        bool counting = reduction is Reduction.Count or Reduction.CountA;
        var numbers = new List<double>();
        int nonEmpty = 0;
        int admitted = 0;
        CellValue first = CellValue.Blank;

        for (int record = 0; record < query.Records; record++)
        {
            if (!query.Admits(record, context))
                continue;

            CellValue value = query.Value(record, column);
            if (value.IsError && !counting)
                return FormulaValue.Scalar(value);

            if (admitted++ == 0)
                first = value;
            if (!value.IsBlank)
                nonEmpty++;
            if (Aggregates.TryNumeric(value, fromRange: true, context, out double number, out _))
                numbers.Add(number);
        }

        return Reduce(reduction, numbers, nonEmpty, admitted, first);
    }

    private static FormulaValue Reduce(
        Reduction reduction, List<double> numbers, int nonEmpty, int admitted, in CellValue first) => reduction switch
        {
            Reduction.Count => FormulaValue.Number(numbers.Count),
            Reduction.CountA => FormulaValue.Number(nonEmpty),
            Reduction.Get => Single(admitted, first),
            Reduction.Sum => FormulaValue.Number(numbers.Sum()),
            Reduction.Product => FormulaValue.Number(Product(numbers)),
            Reduction.Max => FormulaValue.Number(numbers.Count == 0 ? 0 : numbers.Max()),
            Reduction.Min => FormulaValue.Number(numbers.Count == 0 ? 0 : numbers.Min()),
            Reduction.Average => numbers.Count == 0
                ? FormulaValue.Error(ExcelError.DivisionByZero)
                : FormulaValue.Number(numbers.Sum() / numbers.Count),
            _ => Spread(numbers, reduction),
        };

    /// <summary>
    /// <c>DGET</c> asks for the one record that fits. Conditions that fit none are a mistake in
    /// the conditions, conditions that fit several are a mistake in the question, and Excel
    /// tells the two apart.
    /// </summary>
    private static FormulaValue Single(int admitted, in CellValue first) => admitted switch
    {
        0 => FormulaValue.Error(ExcelError.Value),
        1 => FormulaValue.Scalar(first),
        _ => FormulaValue.Error(ExcelError.Number),
    };

    private static double Product(List<double> numbers)
    {
        if (numbers.Count == 0)
            return 0;

        double product = 1;
        foreach (double number in numbers)
            product *= number;

        return product;
    }

    private static FormulaValue Spread(List<double> numbers, Reduction reduction)
    {
        bool sample = reduction is Reduction.Stdev or Reduction.Var;
        int divisor = sample ? numbers.Count - 1 : numbers.Count;
        if (divisor <= 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double mean = numbers.Sum() / numbers.Count;
        double sumOfSquares = 0;
        foreach (double number in numbers)
            sumOfSquares += (number - mean) * (number - mean);

        double variance = sumOfSquares / divisor;
        bool root = reduction is Reduction.Stdev or Reduction.StdevP;
        return FormulaValue.Number(root ? Math.Sqrt(variance) : variance);
    }
}
