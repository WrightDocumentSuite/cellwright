namespace Cellwright.Formulas.Functions;

/// <summary>
/// The paired statistics: correlation, covariance, the straight-line fit read as single
/// numbers, and the hypothesis tests that compare two samples. Pairs where either side is not
/// a number drop out together, which is how Excel walks two ranges with holes.
/// </summary>
internal static class CorrelationFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("CORREL", 2, 2, (c, a) => Paired(c, a, Kind.Correlation), wantsRawReferences: true);
        builder.Add("PEARSON", 2, 2, (c, a) => Paired(c, a, Kind.Correlation), wantsRawReferences: true);
        builder.Add("RSQ", 2, 2, (c, a) => Paired(c, a, Kind.RSquared), wantsRawReferences: true);
        builder.Add("COVAR", 2, 2, (c, a) => Paired(c, a, Kind.CovariancePopulation), wantsRawReferences: true);
        builder.Add("COVARIANCE.P", 2, 2, (c, a) => Paired(c, a, Kind.CovariancePopulation), wantsRawReferences: true);
        builder.Add("COVARIANCE.S", 2, 2, (c, a) => Paired(c, a, Kind.CovarianceSample), wantsRawReferences: true);
        builder.Add("SLOPE", 2, 2, (c, a) => Paired(c, a, Kind.Slope), wantsRawReferences: true);
        builder.Add("INTERCEPT", 2, 2, (c, a) => Paired(c, a, Kind.Intercept), wantsRawReferences: true);
        builder.Add("STEYX", 2, 2, (c, a) => Paired(c, a, Kind.StandardError), wantsRawReferences: true);
        builder.Add("FORECAST", 3, 3, Forecast, wantsRawReferences: true);
        builder.Alias("FORECAST.LINEAR", "FORECAST");
        builder.Add("TTEST", 4, 4, TTest, wantsRawReferences: true);
        builder.Alias("T.TEST", "TTEST");
        builder.Add("FTEST", 2, 2, FTest, wantsRawReferences: true);
        builder.Alias("F.TEST", "FTEST");
        builder.Add("CHITEST", 2, 2, ChiTest, wantsRawReferences: true);
        builder.Alias("CHISQ.TEST", "CHITEST");
        builder.Add("ZTEST", 2, 3, ZTest, wantsRawReferences: true);
        builder.Alias("Z.TEST", "ZTEST");
    }

    private enum Kind
    {
        Correlation,
        RSquared,
        CovariancePopulation,
        CovarianceSample,
        Slope,
        Intercept,
        StandardError,
    }

    /// <summary>The sums every paired statistic is assembled from.</summary>
    private readonly record struct PairedSums(
        int Count, double MeanX, double MeanY, double SumXx, double SumYy, double SumXy);

    private static FormulaValue Paired(EvalContext context, ReadOnlySpan<FormulaValue> args, Kind kind)
    {
        // The observations come first in Excel's signatures (known_y, known_x), so the second
        // argument is the x of the fit.
        if (!TryCollectPairs(context, args[1], args[0], out List<double> xs, out List<double> ys, out CellValue error))
            return FormulaValue.Scalar(error);

        PairedSums sums = Summarize(xs, ys);
        int n = sums.Count;
        double result;
        switch (kind)
        {
            case Kind.Correlation or Kind.RSquared:
                if (n == 0 || sums.SumXx == 0 || sums.SumYy == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                double r = sums.SumXy / Math.Sqrt(sums.SumXx * sums.SumYy);
                result = kind == Kind.Correlation ? r : r * r;
                break;

            case Kind.CovariancePopulation:
                if (n == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                result = sums.SumXy / n;
                break;

            case Kind.CovarianceSample:
                if (n < 2)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                result = sums.SumXy / (n - 1);
                break;

            case Kind.Slope or Kind.Intercept:
                if (n == 0 || sums.SumXx == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                double slope = sums.SumXy / sums.SumXx;
                result = kind == Kind.Slope ? slope : sums.MeanY - (slope * sums.MeanX);
                break;

            default:
                if (n < 3)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                if (sums.SumXx == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                double residual = sums.SumYy - (sums.SumXy * sums.SumXy / sums.SumXx);
                result = Math.Sqrt(Math.Max(residual, 0) / (n - 2));
                break;
        }

        return double.IsNaN(result) || double.IsInfinity(result)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(result);
    }

    private static FormulaValue Forecast(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double at, out CellValue error))
            return FormulaValue.Scalar(error);
        if (!TryCollectPairs(context, args[2], args[1], out List<double> xs, out List<double> ys, out error))
            return FormulaValue.Scalar(error);

        PairedSums sums = Summarize(xs, ys);
        if (sums.Count == 0 || sums.SumXx == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double slope = sums.SumXy / sums.SumXx;
        return FormulaValue.Number(sums.MeanY + (slope * (at - sums.MeanX)));
    }

    /// <summary>
    /// <c>TTEST(a, b, tails, type)</c> — the p-value of Student's t: paired (1), two samples of
    /// one variance (2), or Welch's unequal variances (3).
    /// </summary>
    private static FormulaValue TTest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[2], out int tails, out CellValue error) ||
            !Args.TryInteger(context, args[3], out int type, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (tails is not (1 or 2) || type is not (1 or 2 or 3))
            return FormulaValue.Error(ExcelError.Number);

        double t;
        double freedom;
        if (type == 1)
        {
            if (!TryCollectPairs(context, args[0], args[1], out List<double> xs, out List<double> ys, out error))
                return FormulaValue.Scalar(error);
            int n = xs.Count;
            if (n < 2)
                return FormulaValue.Error(ExcelError.DivisionByZero);

            double meanDifference = 0;
            for (int i = 0; i < n; i++)
                meanDifference += xs[i] - ys[i];
            meanDifference /= n;

            double spread = 0;
            for (int i = 0; i < n; i++)
            {
                double deviation = xs[i] - ys[i] - meanDifference;
                spread += deviation * deviation;
            }

            if (spread == 0)
                return FormulaValue.Error(ExcelError.DivisionByZero);
            t = meanDifference / Math.Sqrt(spread / (n - 1) / n);
            freedom = n - 1;
        }
        else
        {
            (int nx, double meanX, double varianceX) = Sample(context, args[0], out CellValue firstError);
            if (firstError.IsError)
                return FormulaValue.Scalar(firstError);
            (int ny, double meanY, double varianceY) = Sample(context, args[1], out CellValue secondError);
            if (secondError.IsError)
                return FormulaValue.Scalar(secondError);
            if (nx < 2 || ny < 2)
                return FormulaValue.Error(ExcelError.DivisionByZero);

            if (type == 2)
            {
                double pooled = (((nx - 1) * varianceX) + ((ny - 1) * varianceY)) / (nx + ny - 2);
                if (pooled == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                t = (meanX - meanY) / Math.Sqrt(pooled * ((1.0 / nx) + (1.0 / ny)));
                freedom = nx + ny - 2;
            }
            else
            {
                double vx = varianceX / nx;
                double vy = varianceY / ny;
                if (vx + vy == 0)
                    return FormulaValue.Error(ExcelError.DivisionByZero);
                t = (meanX - meanY) / Math.Sqrt(vx + vy);
                freedom = (vx + vy) * (vx + vy) / ((vx * vx / (nx - 1)) + (vy * vy / (ny - 1)));
            }
        }

        double upperTail = StudentUpperTail(Math.Abs(t), freedom);
        return FormulaValue.Number(tails == 1 ? upperTail : 2 * upperTail);
    }

    /// <summary><c>FTEST(a, b)</c> — the two-tailed p of the ratio of the two sample variances.</summary>
    private static FormulaValue FTest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        (int nx, _, double varianceX) = Sample(context, args[0], out CellValue firstError);
        if (firstError.IsError)
            return FormulaValue.Scalar(firstError);
        (int ny, _, double varianceY) = Sample(context, args[1], out CellValue secondError);
        if (secondError.IsError)
            return FormulaValue.Scalar(secondError);
        if (nx < 2 || ny < 2 || varianceX == 0 || varianceY == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double f = varianceX / varianceY;
        double below = FisherBelow(f, nx - 1, ny - 1);
        return FormulaValue.Number(2 * Math.Min(below, 1 - below));
    }

    /// <summary>
    /// <c>CHITEST(actual, expected)</c> — the p of the χ² statistic. Freedom follows the shape:
    /// a true table has (rows−1)(columns−1), a single row or column one fewer than its cells.
    /// </summary>
    private static FormulaValue ChiTest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue actual) ||
            !Operators.TryGetArray(args[1], context, out ArrayValue expected))
        {
            return FormulaValue.Error(ExcelError.Value);
        }

        if (actual.Rows != expected.Rows || actual.Columns != expected.Columns)
            return FormulaValue.Error(ExcelError.NotAvailable);

        double statistic = 0;
        for (int i = 0; i < actual.Count; i++)
        {
            if (!Aggregates.TryNumeric(actual.Values[i], fromRange: true, context, out double a, out CellValue error) ||
                !Aggregates.TryNumeric(expected.Values[i], fromRange: true, context, out double e, out error))
            {
                if (error.IsError)
                    return FormulaValue.Scalar(error);
                continue;
            }

            if (e == 0)
                return FormulaValue.Error(ExcelError.DivisionByZero);
            statistic += (a - e) * (a - e) / e;
        }

        double freedom = actual.Rows > 1 && actual.Columns > 1
            ? (actual.Rows - 1) * (double)(actual.Columns - 1)
            : Math.Max(actual.Count - 1, 1);
        return FormulaValue.Number(SpecialFunctions.UpperGamma(freedom / 2, statistic / 2));
    }

    /// <summary>
    /// <c>ZTEST(array, x, [sigma])</c> — the one-tailed p that the sample mean sits above
    /// <c>x</c>, with the population deviation given or taken from the sample.
    /// </summary>
    private static FormulaValue ZTest(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        (int n, double mean, double variance) = Sample(context, args[0], out CellValue error);
        if (error.IsError)
            return FormulaValue.Scalar(error);
        if (!Args.TryNumber(context, args[1], out double x, out CellValue bad))
            return FormulaValue.Scalar(bad);

        double sigma;
        if (args.Length > 2 && !args[2].IsBlank)
        {
            if (!Args.TryNumber(context, args[2], out sigma, out bad))
                return FormulaValue.Scalar(bad);
        }
        else
        {
            if (n < 2)
                return FormulaValue.Error(ExcelError.DivisionByZero);
            sigma = Math.Sqrt(variance);
        }

        if (n == 0 || sigma <= 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        double z = (mean - x) / (sigma / Math.Sqrt(n));
        return FormulaValue.Number(1 - SpecialFunctions.NormalBelow(z));
    }

    /// <summary>P(T &gt; t) for Student's t with the given freedom.</summary>
    private static double StudentUpperTail(double t, double freedom) =>
        SpecialFunctions.IncompleteBeta(freedom / 2, 0.5, freedom / (freedom + (t * t))) / 2;

    /// <summary>P(F ≤ f) for Fisher's F with the given freedoms.</summary>
    private static double FisherBelow(double f, double d1, double d2) =>
        SpecialFunctions.IncompleteBeta(d1 / 2, d2 / 2, d1 * f / ((d1 * f) + d2));

    /// <summary>Count, mean and sample variance of one argument's numbers.</summary>
    private static (int Count, double Mean, double Variance) Sample(
        EvalContext context, in FormulaValue argument, out CellValue error)
    {
        List<double> values = [];
        error = ExtendedStatisticalFunctions.CollectNumbers(context, argument, values);
        if (error.IsError || values.Count == 0)
            return (values.Count, 0, 0);

        double mean = 0;
        foreach (double value in values)
            mean += value;
        mean /= values.Count;

        double spread = 0;
        foreach (double value in values)
            spread += (value - mean) * (value - mean);

        return (values.Count, mean, values.Count > 1 ? spread / (values.Count - 1) : 0);
    }

    /// <summary>
    /// Walks two equally shaped arguments, keeping the positions where both sides are numbers.
    /// </summary>
    internal static bool TryCollectPairs(
        EvalContext context, in FormulaValue first, in FormulaValue second,
        out List<double> xs, out List<double> ys, out CellValue error)
    {
        xs = [];
        ys = [];
        error = CellValue.Blank;
        if (!Operators.TryGetArray(first, context, out ArrayValue left) ||
            !Operators.TryGetArray(second, context, out ArrayValue right))
        {
            error = CellValue.Error(ExcelError.Value);
            return false;
        }

        if (left.Count != right.Count)
        {
            error = CellValue.Error(ExcelError.NotAvailable);
            return false;
        }

        for (int i = 0; i < left.Count; i++)
        {
            bool hasX = Aggregates.TryNumeric(left.Values[i], fromRange: true, context, out double x, out CellValue bad);
            if (bad.IsError)
            {
                error = bad;
                return false;
            }

            bool hasY = Aggregates.TryNumeric(right.Values[i], fromRange: true, context, out double y, out bad);
            if (bad.IsError)
            {
                error = bad;
                return false;
            }

            if (!hasX || !hasY)
                continue;

            xs.Add(x);
            ys.Add(y);
        }

        return true;
    }

    /// <summary>The centered sums of a pair list.</summary>
    private static PairedSums Summarize(List<double> xs, List<double> ys)
    {
        int n = xs.Count;
        if (n == 0)
            return default;

        double meanX = 0;
        double meanY = 0;
        for (int i = 0; i < n; i++)
        {
            meanX += xs[i];
            meanY += ys[i];
        }

        meanX /= n;
        meanY /= n;

        double sumXx = 0;
        double sumYy = 0;
        double sumXy = 0;
        for (int i = 0; i < n; i++)
        {
            double dx = xs[i] - meanX;
            double dy = ys[i] - meanY;
            sumXx += dx * dx;
            sumYy += dy * dy;
            sumXy += dx * dy;
        }

        return new PairedSums(n, meanX, meanY, sumXx, sumYy, sumXy);
    }
}
