namespace Cellwright.Formulas.Functions;

/// <summary>
/// The hyperbolic curves, the combinatorial counts and the paired-array sums of §18.17.7 —
/// small functions the catalog was missing not for difficulty but for number.
/// </summary>
internal static class ExtendedMathFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SINH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Sinh));
        builder.Add("COSH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Cosh));
        builder.Add("TANH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Tanh));
        builder.Add("ASINH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Asinh));
        builder.Add("ACOSH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Acosh));
        builder.Add("ATANH", 1, 1, (c, a) => Args.Map(c, a[0], Math.Atanh));
        builder.Add("SEC", 1, 1, (c, a) => Args.Map(c, a[0], static x => 1 / Math.Cos(x)));
        builder.Add("CSC", 1, 1, (c, a) => Args.Map(c, a[0], static x => 1 / Math.Sin(x)));
        builder.Add("COT", 1, 1, (c, a) => Args.Map(c, a[0], static x => Math.Cos(x) / Math.Sin(x)));
        builder.Add("SECH", 1, 1, (c, a) => Args.Map(c, a[0], static x => 1 / Math.Cosh(x)));
        builder.Add("CSCH", 1, 1, (c, a) => Args.Map(c, a[0], static x => 1 / Math.Sinh(x)));
        builder.Add("COTH", 1, 1, (c, a) => Args.Map(c, a[0], static x => Math.Cosh(x) / Math.Sinh(x)));
        builder.Add("ACOT", 1, 1, (c, a) => Args.Map(c, a[0], static x => (Math.PI / 2) - Math.Atan(x)));
        builder.Add("ACOTH", 1, 1, (c, a) => Args.Map(c, a[0], static x => Math.Atanh(1 / x)));
        builder.Add("COMBIN", 2, 2, Combin);
        builder.Add("PERMUT", 2, 2, Permut);
        builder.Add("MROUND", 2, 2, MRound);
        builder.Add("SQRTPI", 1, 1, (c, a) => Args.Map(c, a[0], static x => Math.Sqrt(x * Math.PI)));
        builder.Add("FACTDOUBLE", 1, 1, FactDouble);
        builder.Add("MULTINOMIAL", 1, FunctionDefinition.Unlimited, Multinomial, wantsRawReferences: true);
        builder.Add("SERIESSUM", 4, 4, SeriesSum, wantsRawReferences: true);
        builder.Add("SUMSQ", 1, FunctionDefinition.Unlimited, SumSquares, wantsRawReferences: true);
        builder.Add("SUMX2MY2", 2, 2, (c, a) => PairedSum(c, a, static (x, y) => (x * x) - (y * y)), wantsRawReferences: true);
        builder.Add("SUMX2PY2", 2, 2, (c, a) => PairedSum(c, a, static (x, y) => (x * x) + (y * y)), wantsRawReferences: true);
        builder.Add("SUMXMY2", 2, 2, (c, a) => PairedSum(c, a, static (x, y) => (x - y) * (x - y)), wantsRawReferences: true);
    }

    /// <summary>
    /// <c>COMBIN(n, k)</c> — ways to choose <c>k</c> of <c>n</c> without order. Computed as a
    /// running product to stay exact far past what factorials of doubles would allow.
    /// </summary>
    private static FormulaValue Combin(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double nRaw, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double kRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double n = Math.Truncate(nRaw);
        double k = Math.Truncate(kRaw);
        if (n < 0 || k < 0 || n < k)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(Choose(n, k));
    }

    /// <summary>Binomial coefficient by the smaller side, shared with the statistics catalog.</summary>
    internal static double Choose(double n, double k)
    {
        k = Math.Min(k, n - k);
        double result = 1;
        for (int i = 1; i <= k; i++)
            result = result * (n - k + i) / i;
        return Math.Round(result);
    }

    private static FormulaValue Permut(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double nRaw, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double kRaw, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double n = Math.Truncate(nRaw);
        double k = Math.Truncate(kRaw);
        if (n < 0 || k < 0 || n < k)
            return FormulaValue.Error(ExcelError.Number);

        double result = 1;
        for (double i = n - k + 1; i <= n; i++)
            result *= i;
        return FormulaValue.Number(result);
    }

    /// <summary>
    /// <c>MROUND(x, multiple)</c> — the nearest multiple, halves away from zero. The two
    /// arguments must pull the same way; opposite signs are refused like CEILING refuses them.
    /// </summary>
    private static FormulaValue MRound(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double multiple, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (multiple == 0 || number == 0)
            return FormulaValue.Number(0);
        if (Math.Sign(number) != Math.Sign(multiple))
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(Math.Round(number / multiple, MidpointRounding.AwayFromZero) * multiple);
    }

    private static FormulaValue FactDouble(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double raw, out CellValue error))
            return FormulaValue.Scalar(error);

        double n = Math.Truncate(raw);
        if (n < 0)
            return FormulaValue.Error(ExcelError.Number);

        double result = 1;
        for (double i = n; i > 1; i -= 2)
            result *= i;
        return double.IsInfinity(result) ? FormulaValue.Error(ExcelError.Number) : FormulaValue.Number(result);
    }

    /// <summary>
    /// <c>MULTINOMIAL(…)</c> — the ratio of the factorial of the sum to the product of the
    /// factorials, built incrementally as a product of binomials so it does not overflow on
    /// numbers whose factorials would.
    /// </summary>
    private static FormulaValue Multinomial(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        double result = 1;
        double runningSum = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!Aggregates.TryNumeric(value, arg.Kind != FormulaValueKind.Scalar, context, out double number, out CellValue error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                double count = Math.Truncate(number);
                if (count < 0)
                    return FormulaValue.Error(ExcelError.Number);

                runningSum += count;
                result *= Choose(runningSum, count);
            }
        }

        return double.IsInfinity(result) ? FormulaValue.Error(ExcelError.Number) : FormulaValue.Number(result);
    }

    /// <summary>
    /// <c>SERIESSUM(x, n, m, coefficients)</c> — the power series
    /// <c>a₁xⁿ + a₂xⁿ⁺ᵐ + a₃xⁿ⁺²ᵐ + …</c> with the coefficients in reading order.
    /// </summary>
    private static FormulaValue SeriesSum(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double n, out error) ||
            !Args.TryNumber(context, args[2], out double m, out error))
        {
            return FormulaValue.Scalar(error);
        }

        double total = 0;
        int index = 0;
        foreach (CellValue value in new ArgumentValues(args[3], context))
        {
            if (!Aggregates.TryNumeric(value, args[3].Kind != FormulaValueKind.Scalar, context, out double coefficient, out CellValue bad))
            {
                if (bad.IsError)
                    return FormulaValue.Scalar(bad);
                continue;
            }

            total += coefficient * Math.Pow(x, n + (index * m));
            index++;
        }

        return double.IsNaN(total) || double.IsInfinity(total)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(total);
    }

    private static FormulaValue SumSquares(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        double total = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!Aggregates.TryNumeric(value, arg.Kind != FormulaValueKind.Scalar, context, out double number, out CellValue error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                total += number * number;
            }
        }

        return FormulaValue.Number(total);
    }

    /// <summary>
    /// The three paired-array sums. The arrays must agree in size; a pair where either side is
    /// not a number contributes nothing, which is how Excel skips text and blanks in them.
    /// </summary>
    private static FormulaValue PairedSum(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Func<double, double, double> term)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue first) ||
            !Operators.TryGetArray(args[1], context, out ArrayValue second))
        {
            return FormulaValue.Error(ExcelError.Value);
        }

        if (first.Count != second.Count)
            return FormulaValue.Error(ExcelError.NotAvailable);

        double total = 0;
        for (int i = 0; i < first.Count; i++)
        {
            bool hasX = Aggregates.TryNumeric(first.Values[i], fromRange: true, context, out double x, out CellValue error);
            if (error.IsError)
                return FormulaValue.Scalar(error);
            bool hasY = Aggregates.TryNumeric(second.Values[i], fromRange: true, context, out double y, out error);
            if (error.IsError)
                return FormulaValue.Scalar(error);
            if (!hasX || !hasY)
                continue;

            total += term(x, y);
        }

        return FormulaValue.Number(total);
    }
}
