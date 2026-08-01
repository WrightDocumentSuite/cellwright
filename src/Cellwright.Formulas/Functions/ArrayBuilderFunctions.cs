namespace Cellwright.Formulas.Functions;

/// <summary>
/// The two dynamic array functions that make an array out of nothing rather than reshaping one
/// they were handed.
/// </summary>
internal static class ArrayBuilderFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SEQUENCE", 1, 4, Sequence);
        builder.Add("RANDARRAY", 0, 5, RandomArray, isVolatile: true);
    }

    /// <summary>
    /// <c>SEQUENCE(rows, [columns], [start], [step])</c>: the counting numbers laid out in a
    /// rectangle, filled across each row before moving down.
    /// </summary>
    private static FormulaValue Sequence(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int rows, out CellValue error))
            return FormulaValue.Scalar(error);

        int columns = 1;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryInteger(context, args[1], out columns, out error))
            return FormulaValue.Scalar(error);

        double start = 1;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryNumber(context, args[2], out start, out error))
            return FormulaValue.Scalar(error);

        double step = 1;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryNumber(context, args[3], out step, out error))
            return FormulaValue.Scalar(error);

        if (!TrySize(rows, columns, out FormulaValue refused))
            return refused;

        var values = new CellValue[rows * columns];
        for (int i = 0; i < values.Length; i++)
            values[i] = CellValue.Number(start + (i * step));

        return FormulaValue.Array(new ArrayValue(rows, columns, values));
    }

    /// <summary>
    /// <c>RANDARRAY([rows], [columns], [min], [max], [whole])</c>: a rectangle of random
    /// numbers, whole ones between the bounds inclusive or fractions from the interval below the top.
    /// </summary>
    private static FormulaValue RandomArray(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        int rows = 1;
        if (args.Length > 0 && !args[0].IsBlank && !Args.TryInteger(context, args[0], out rows, out CellValue error))
            return FormulaValue.Scalar(error);

        int columns = 1;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryInteger(context, args[1], out columns, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        double minimum = 0;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryNumber(context, args[2], out minimum, out CellValue refused))
            return FormulaValue.Scalar(refused);

        double maximum = 1;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryNumber(context, args[3], out maximum, out CellValue denied))
            return FormulaValue.Scalar(denied);

        bool whole = false;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryBoolean(context, args[4], out whole, out CellValue barred))
            return FormulaValue.Scalar(barred);

        if (minimum > maximum)
            return FormulaValue.Error(ExcelError.Value);
        if (!TrySize(rows, columns, out FormulaValue failure))
            return failure;

        return Fill(rows, columns, minimum, maximum, whole);
    }

    private static FormulaValue Fill(int rows, int columns, double minimum, double maximum, bool whole)
    {
        long low = (long)Math.Ceiling(minimum);
        long high = (long)Math.Floor(maximum);

        // Asked for whole numbers with no whole number between the bounds, there is nothing to
        // hand back — a fraction would not be an answer to the question.
        if (whole && low > high)
            return FormulaValue.Error(ExcelError.Value);

        var values = new CellValue[rows * columns];
        for (int i = 0; i < values.Length; i++)
        {
            values[i] = CellValue.Number(whole
                ? Random.Shared.NextInt64(low, high + 1)
                : minimum + (Random.Shared.NextDouble() * (maximum - minimum)));
        }

        return FormulaValue.Array(new ArrayValue(rows, columns, values));
    }

    /// <summary>
    /// Checks a requested shape. An empty rectangle is <c>#CALC!</c> — a shape with nothing in
    /// it, which is a different complaint from a shape that makes no sense at all.
    /// </summary>
    private static bool TrySize(int rows, int columns, out FormulaValue refused)
    {
        refused = FormulaValue.Blank;
        if (rows < 0 || columns < 0)
            refused = FormulaValue.Error(ExcelError.Value);
        else if (rows == 0 || columns == 0)
            refused = FormulaValue.Error(ExcelError.Calc);
        else if ((long)rows * columns > ArrayValue.MaxCells)
            refused = FormulaValue.Error(ExcelError.Number);

        return refused.IsBlank;
    }
}
