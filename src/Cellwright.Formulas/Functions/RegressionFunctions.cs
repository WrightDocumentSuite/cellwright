namespace Cellwright.Formulas.Functions;

/// <summary>
/// <c>LINEST</c>: the straight line, or plane, that comes closest to a set of observations.
/// </summary>
/// <remarks>
/// The fit is solved through the normal equations rather than a decomposition. That is the less
/// careful of the two methods on badly conditioned data, but it hands back the inverse of the
/// cross-product matrix as a by-product, and the whole second row of the statistics block is
/// read straight off its diagonal.
/// </remarks>
internal static class RegressionFunctions
{
    public static void Register(FunctionRegistryBuilder builder) =>
        builder.Add("LINEST", 1, 4, Linest, wantsRawReferences: true);

    private static FormulaValue Linest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue observed))
            return FormulaValue.Error(ExcelError.Value);

        if (!TryReadDesign(context, args, observed, out double[] y, out double[,] x, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        bool intercept = true;
        if (args.Length > 2 && !args[2].IsBlank &&
            !Args.TryBoolean(context, args[2], out intercept, out CellValue error))
        {
            return FormulaValue.Scalar(error);
        }

        bool statistics = false;
        if (args.Length > 3 && !args[3].IsBlank &&
            !Args.TryBoolean(context, args[3], out statistics, out CellValue failure))
        {
            return FormulaValue.Scalar(failure);
        }

        return Fit(y, x, intercept, statistics);
    }

    /// <summary>
    /// Reads the observations and the columns that explain them. The pair may be written either
    /// way up — a column of observations beside columns of variables, or a row beside rows — and
    /// which it is follows from the shape of the observations. Shared with TREND and GROWTH,
    /// whose designs read the same way.
    /// </summary>
    internal static bool TryReadDesign(
        EvalContext context, ReadOnlySpan<FormulaValue> args, ArrayValue observed,
        out double[] y, out double[,] x, out CellValue rejected)
    {
        x = new double[0, 0];
        rejected = CellValue.Error(ExcelError.Reference);
        int n = observed.Count;
        y = new double[n];
        for (int i = 0; i < n; i++)
        {
            if (!Numeric(observed.Values[i], context, out y[i]))
            {
                rejected = CellValue.Error(ExcelError.Value);
                return false;
            }
        }

        if (n == 0)
            return false;

        // Left out, the variable is simply the position of each observation, which is what makes
        // LINEST(y) a trend over time.
        if (args.Length < 2 || args[1].IsBlank)
        {
            x = new double[n, 1];
            for (int i = 0; i < n; i++)
                x[i, 0] = i + 1;
            return true;
        }

        if (!Operators.TryGetArray(args[1], context, out ArrayValue explanatory))
            return false;

        bool down = observed.Rows >= observed.Columns;
        int variables = explanatory.Count == n ? 1 : down ? explanatory.Columns : explanatory.Rows;
        int rows = explanatory.Count == n ? n : down ? explanatory.Rows : explanatory.Columns;
        if (rows != n || variables < 1)
            return false;

        return TryFill(context, explanatory, n, variables, down, ref x, ref rejected);
    }

    private static bool TryFill(
        EvalContext context, ArrayValue explanatory, int n, int variables, bool down,
        ref double[,] x, ref CellValue rejected)
    {
        x = new double[n, variables];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < variables; j++)
            {
                CellValue cell = variables == 1 ? explanatory.Values[i] : down ? explanatory[i, j] : explanatory[j, i];
                if (!Numeric(cell, context, out x[i, j]))
                {
                    rejected = CellValue.Error(ExcelError.Value);
                    return false;
                }
            }
        }

        return true;
    }

    private static bool Numeric(in CellValue value, EvalContext context, out double number) =>
        Aggregates.TryNumeric(value, fromRange: true, context, out number, out _);

    private static FormulaValue Fit(double[] y, double[,] x, bool intercept, bool statistics)
    {
        if (!TrySolve(y, x, intercept, out double[] coefficients, out double[,] inverse))
            return FormulaValue.Error(ExcelError.Number);

        int variables = x.GetLength(1);
        return FormulaValue.Array(statistics
            ? WithStatistics(y, x, coefficients, inverse, intercept)
            : Coefficients(coefficients, variables, intercept));
    }

    /// <summary>
    /// Solves the least-squares fit through the normal equations. The coefficients arrive in
    /// term order — the variables first, then the constant when asked for one.
    /// </summary>
    internal static bool TrySolve(
        double[] y, double[,] x, bool intercept, out double[] coefficients, out double[,] inverse)
    {
        int n = y.Length;
        int variables = x.GetLength(1);
        int terms = variables + (intercept ? 1 : 0);
        coefficients = [];
        inverse = new double[0, 0];
        if (n < terms)
            return false;

        var normal = new double[terms, terms];
        var moment = new double[terms];
        for (int i = 0; i < n; i++)
        {
            for (int a = 0; a < terms; a++)
            {
                double left = Term(x, i, a, variables);
                moment[a] += left * y[i];
                for (int b = 0; b < terms; b++)
                    normal[a, b] += left * Term(x, i, b, variables);
            }
        }

        if (!Matrix.TryInvert(normal, out inverse))
            return false;

        coefficients = new double[terms];
        for (int a = 0; a < terms; a++)
        {
            for (int b = 0; b < terms; b++)
                coefficients[a] += inverse[a, b] * moment[b];
        }

        return true;
    }

    /// <summary>One term of the fit at one observation; the last term is the constant.</summary>
    internal static double Term(double[,] x, int row, int term, int variables) =>
        term < variables ? x[row, term] : 1;

    /// <summary>
    /// The coefficients as Excel lays them out: the variables in reverse, then the constant.
    /// A fit told to pass through the origin still reports the constant, as a zero.
    /// </summary>
    private static ArrayValue Coefficients(double[] coefficients, int variables, bool intercept)
    {
        var cells = new CellValue[variables + 1];
        for (int j = 0; j < variables; j++)
            cells[j] = CellValue.Number(coefficients[variables - 1 - j]);

        cells[variables] = CellValue.Number(intercept ? coefficients[variables] : 0);
        return new ArrayValue(1, variables + 1, cells);
    }

    /// <summary>
    /// The five-row block: the coefficients, the uncertainty in each, then how much of the
    /// variation the fit accounts for and how confident that makes it. Cells the shape leaves
    /// over are <c>#N/A</c>, which is how Excel fills them too.
    /// </summary>
    private static ArrayValue WithStatistics(
        double[] y, double[,] x, double[] coefficients, double[,] inverse, bool intercept)
    {
        int variables = x.GetLength(1);
        int columns = variables + 1;
        (double residual, double total) = Spread(y, x, coefficients, intercept);
        int freedom = y.Length - coefficients.Length;
        double regression = total - residual;
        double deviation = freedom > 0 ? Math.Sqrt(residual / freedom) : double.NaN;

        var cells = new CellValue[5 * columns];
        Array.Fill(cells, CellValue.Error(ExcelError.NotAvailable));
        for (int j = 0; j <= variables; j++)
        {
            // The variables are listed in reverse and the constant last, so the column a term
            // lands in is not the term's own position.
            bool constant = j == variables;
            int term = constant ? variables : variables - 1 - j;
            cells[j] = CellValue.Number(constant && !intercept ? 0 : coefficients[term]);
            cells[columns + j] = constant && !intercept
                ? CellValue.Error(ExcelError.NotAvailable)
                : Statistic(deviation * Math.Sqrt(inverse[term, term]));
        }

        cells[(2 * columns) + 0] = Statistic(total == 0 ? double.NaN : regression / total);
        cells[(2 * columns) + 1] = Statistic(deviation);
        cells[(3 * columns) + 0] = Statistic(residual == 0 ? double.NaN : regression / variables / (residual / freedom));
        cells[(3 * columns) + 1] = CellValue.Number(freedom);
        cells[(4 * columns) + 0] = CellValue.Number(regression);
        cells[(4 * columns) + 1] = CellValue.Number(residual);
        return new ArrayValue(5, columns, cells);
    }

    /// <summary>
    /// What the fit failed to account for, and what there was to account for in the first place.
    /// A fit through the origin measures the second against zero rather than against the mean.
    /// </summary>
    private static (double Residual, double Total) Spread(
        double[] y, double[,] x, double[] coefficients, bool intercept)
    {
        int variables = x.GetLength(1);
        double mean = 0;
        if (intercept)
        {
            foreach (double value in y)
                mean += value;
            mean /= y.Length;
        }

        double residual = 0;
        double total = 0;
        for (int i = 0; i < y.Length; i++)
        {
            double predicted = 0;
            for (int a = 0; a < coefficients.Length; a++)
                predicted += coefficients[a] * Term(x, i, a, variables);

            residual += (y[i] - predicted) * (y[i] - predicted);
            total += (y[i] - mean) * (y[i] - mean);
        }

        return (residual, total);
    }

    private static CellValue Statistic(double value) =>
        double.IsNaN(value) || double.IsInfinity(value)
            ? CellValue.Error(ExcelError.Number)
            : CellValue.Number(value);
}
