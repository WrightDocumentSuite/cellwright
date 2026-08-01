namespace Cellwright.Formulas.Functions;

/// <summary>
/// The fits that answer with values rather than coefficients — <c>TREND</c> and <c>GROWTH</c> —
/// and <c>LOGEST</c>, the exponential twin of LINEST. All three ride the LINEST machinery;
/// the exponential pair only takes the logarithm on the way in and undoes it on the way out.
/// </summary>
internal static class TrendFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("TREND", 1, 4, (c, a) => Predict(c, a, exponential: false), wantsRawReferences: true);
        builder.Add("GROWTH", 1, 4, (c, a) => Predict(c, a, exponential: true), wantsRawReferences: true);
        builder.Add("LOGEST", 1, 4, Logest, wantsRawReferences: true);
    }

    private static FormulaValue Predict(EvalContext context, ReadOnlySpan<FormulaValue> args, bool exponential)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue observed))
            return FormulaValue.Error(ExcelError.Value);
        if (!RegressionFunctions.TryReadDesign(context, args, observed, out double[] y, out double[,] x, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        bool intercept = true;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryBoolean(context, args[3], out intercept, out CellValue error))
            return FormulaValue.Scalar(error);

        if (exponential && !TryLogarithms(y))
            return FormulaValue.Error(ExcelError.Number);
        if (!RegressionFunctions.TrySolve(y, x, intercept, out double[] coefficients, out _))
            return FormulaValue.Error(ExcelError.Number);

        int variables = x.GetLength(1);
        if (!TryReadTargets(context, args, x, variables, out double[,] targets, out int rows, out int columns))
            return FormulaValue.Error(ExcelError.Value);

        var cells = new CellValue[rows * columns];
        for (int i = 0; i < cells.Length; i++)
        {
            double predicted = 0;
            for (int a = 0; a < coefficients.Length; a++)
                predicted += coefficients[a] * RegressionFunctions.Term(targets, i, a, variables);
            if (exponential)
                predicted = Math.Exp(predicted);
            cells[i] = double.IsNaN(predicted) || double.IsInfinity(predicted)
                ? CellValue.Error(ExcelError.Number)
                : CellValue.Number(predicted);
        }

        return FormulaValue.Array(new ArrayValue(rows, columns, cells));
    }

    /// <summary>
    /// <c>LOGEST</c> — LINEST on the logarithms, with the first row exponentiated so that the
    /// answer reads <c>y = b·m₁ˣ</c>. The statistics rows describe the logarithmic fit, which
    /// is what Excel reports too.
    /// </summary>
    private static FormulaValue Logest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue observed))
            return FormulaValue.Error(ExcelError.Value);
        if (!RegressionFunctions.TryReadDesign(context, args, observed, out double[] y, out double[,] x, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        bool intercept = true;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryBoolean(context, args[2], out intercept, out CellValue error))
            return FormulaValue.Scalar(error);
        bool statistics = false;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryBoolean(context, args[3], out statistics, out CellValue failure))
            return FormulaValue.Scalar(failure);

        if (!TryLogarithms(y))
            return FormulaValue.Error(ExcelError.Number);

        // The linear fit of the logarithms is asked from LINEST itself, then the top row is
        // brought back from the log scale.
        FormulaValue linear = LinestOnLogs(context, args, y, x, intercept, statistics);
        if (linear.Kind != FormulaValueKind.Array)
            return linear;

        ArrayValue block = linear.AsArray;
        var cells = (CellValue[])block.Values.Clone();
        for (int j = 0; j < block.Columns; j++)
        {
            if (cells[j].Kind == CellValueKind.Number)
                cells[j] = CellValue.Number(Math.Exp(cells[j].GetNumber()));
        }

        return FormulaValue.Array(new ArrayValue(block.Rows, block.Columns, cells));
    }

    /// <summary>Asks LINEST itself for the fit of the logarithms, statistics block and all.</summary>
    private static FormulaValue LinestOnLogs(
        EvalContext context, ReadOnlySpan<FormulaValue> args, double[] y, double[,] x, bool intercept, bool statistics)
    {
        _ = x;
        var forwarded = new FormulaValue[4];
        forwarded[0] = FormulaValue.Array(Rewrap(y));
        forwarded[1] = args.Length > 1 ? args[1] : FormulaValue.Blank;
        forwarded[2] = FormulaValue.Boolean(intercept);
        forwarded[3] = FormulaValue.Boolean(statistics);
        FunctionDefinition linest = FunctionRegistry.Resolve("LINEST")!;
        return linest.Handler(context, forwarded);
    }

    private static ArrayValue Rewrap(double[] y)
    {
        var cells = new CellValue[y.Length];
        for (int i = 0; i < y.Length; i++)
            cells[i] = CellValue.Number(y[i]);
        return new ArrayValue(y.Length, 1, cells);
    }

    private static bool TryLogarithms(double[] y)
    {
        for (int i = 0; i < y.Length; i++)
        {
            if (y[i] <= 0)
                return false;
            y[i] = Math.Log(y[i]);
        }

        return true;
    }

    /// <summary>
    /// The points to predict at: the third argument when given, the fit's own points otherwise.
    /// A single-variable target keeps its shape; a multi-variable one reads like the design.
    /// </summary>
    private static bool TryReadTargets(
        EvalContext context, ReadOnlySpan<FormulaValue> args, double[,] x, int variables,
        out double[,] targets, out int rows, out int columns)
    {
        targets = x;
        rows = x.GetLength(0);
        columns = 1;
        if (args.Length < 3 || args[2].IsBlank)
            return true;

        if (!Operators.TryGetArray(args[2], context, out ArrayValue wanted))
            return false;

        if (variables == 1)
        {
            rows = wanted.Rows;
            columns = wanted.Columns;
            targets = new double[wanted.Count, 1];
            for (int i = 0; i < wanted.Count; i++)
            {
                if (!Aggregates.TryNumeric(wanted.Values[i], fromRange: true, context, out targets[i, 0], out _))
                    return false;
            }

            return true;
        }

        bool down = wanted.Columns == variables;
        rows = down ? wanted.Rows : wanted.Columns;
        columns = 1;
        targets = new double[rows, variables];
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < variables; j++)
            {
                CellValue cell = down ? wanted[i, j] : wanted[j, i];
                if (!Aggregates.TryNumeric(cell, fromRange: true, context, out targets[i, j], out _))
                    return false;
            }
        }

        return true;
    }
}
