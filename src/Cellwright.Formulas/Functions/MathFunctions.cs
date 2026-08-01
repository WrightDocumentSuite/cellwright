namespace Cellwright.Formulas.Functions;

/// <summary>
/// Arithmetic and rounding functions.
/// </summary>
internal static class MathFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SUM", 1, FunctionDefinition.Unlimited, Sum, wantsRawReferences: true);
        builder.Add("PRODUCT", 1, FunctionDefinition.Unlimited, Product, wantsRawReferences: true);
        builder.Add("ABS", 1, 1, (c, a) => Args.Map(c, a[0], Math.Abs));
        builder.Add("SIGN", 1, 1, (c, a) => Args.Map(c, a[0], x => Math.Sign(x)));
        builder.Add("SQRT", 1, 1, (c, a) => Args.Map(c, a[0], x => x < 0 ? double.NaN : Math.Sqrt(x)));
        builder.Add("POWER", 2, 2, Power);
        builder.Add("EXP", 1, 1, (c, a) => Args.Map(c, a[0], Math.Exp));
        builder.Add("LN", 1, 1, (c, a) => Args.Map(c, a[0], x => x <= 0 ? double.NaN : Math.Log(x)));
        builder.Add("LOG10", 1, 1, (c, a) => Args.Map(c, a[0], x => x <= 0 ? double.NaN : Math.Log10(x)));
        builder.Add("LOG", 1, 2, Log);
        builder.Add("MOD", 2, 2, Mod);
        builder.Add("QUOTIENT", 2, 2, Quotient);
        builder.Add("INT", 1, 1, (c, a) => Args.Map(c, a[0], Math.Floor));
        builder.Add("TRUNC", 1, 2, (c, a) => RoundLike(c, a, static (x, d) => Math.Truncate(x * Math.Pow(10, d)) / Math.Pow(10, d)));
        builder.Add("ROUND", 1, 2, (c, a) => RoundLike(c, a, Args.Round));
        builder.Add("ROUNDUP", 1, 2, (c, a) => RoundLike(c, a, static (x, d) => Away(x, d, up: true)));
        builder.Add("ROUNDDOWN", 1, 2, (c, a) => RoundLike(c, a, static (x, d) => Away(x, d, up: false)));
        builder.Add("CEILING", 2, 2, (c, a) => Step(c, a, ceiling: true));
        builder.Add("FLOOR", 2, 2, (c, a) => Step(c, a, ceiling: false));
        builder.Add("PI", 0, 0, static (_, _) => FormulaValue.Number(Math.PI));
        builder.Add("RAND", 0, 0, static (_, _) => FormulaValue.Number(Random.Shared.NextDouble()), isVolatile: true);
        builder.Add("RANDBETWEEN", 2, 2, RandBetween, isVolatile: true);
        builder.Add("SIN", 1, 1, (c, a) => Args.Map(c, a[0], Math.Sin));
        builder.Add("COS", 1, 1, (c, a) => Args.Map(c, a[0], Math.Cos));
        builder.Add("TAN", 1, 1, (c, a) => Args.Map(c, a[0], Math.Tan));
        builder.Add("ASIN", 1, 1, (c, a) => Args.Map(c, a[0], Math.Asin));
        builder.Add("ACOS", 1, 1, (c, a) => Args.Map(c, a[0], Math.Acos));
        builder.Add("ATAN", 1, 1, (c, a) => Args.Map(c, a[0], Math.Atan));
        builder.Add("ATAN2", 2, 2, Atan2);
        builder.Add("DEGREES", 1, 1, (c, a) => Args.Map(c, a[0], x => x * 180 / Math.PI));
        builder.Add("RADIANS", 1, 1, (c, a) => Args.Map(c, a[0], x => x * Math.PI / 180));
        builder.Add("EVEN", 1, 1, (c, a) => Args.Map(c, a[0], x => Parity(x, even: true)));
        builder.Add("ODD", 1, 1, (c, a) => Args.Map(c, a[0], x => Parity(x, even: false)));
        builder.Add("FACT", 1, 1, Factorial);
        builder.Add("GCD", 1, FunctionDefinition.Unlimited, (c, a) => IntegerFold(c, a, Gcd), wantsRawReferences: true);
        builder.Add("LCM", 1, FunctionDefinition.Unlimited, (c, a) => IntegerFold(c, a, Lcm), wantsRawReferences: true);
    }

    private static FormulaValue Sum(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        double total = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!Aggregates.TryNumeric(value, fromRange: arg.Kind != FormulaValueKind.Scalar, context, out double number, out CellValue error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                total += number;
            }
        }

        return FormulaValue.Number(total);
    }

    private static FormulaValue Product(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        double product = 1;
        bool any = false;
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

                product *= number;
                any = true;
            }
        }

        return FormulaValue.Number(any ? product : 0);
    }

    private static FormulaValue Power(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double y, out error))
        {
            return FormulaValue.Scalar(error);
        }

        // Zero to the zero is one in mathematics and in .NET, and #NUM! in Excel
        // (MS-OI29500 2.1.1018).
        if (x == 0 && y == 0)
            return FormulaValue.Error(ExcelError.Number);

        double result = Math.Pow(x, y);
        return double.IsNaN(result) ? FormulaValue.Error(ExcelError.Number) : FormulaValue.Number(result);
    }

    private static FormulaValue Log(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error))
            return FormulaValue.Scalar(error);

        double logBase = 10;
        if (args.Length > 1 && !Args.TryNumber(context, args[1], out logBase, out error))
            return FormulaValue.Scalar(error);

        // A base of one has no logarithm because the division by log(1) is a division by zero,
        // and that is the error Excel reports rather than #NUM! (MS-OI29500 2.1.986).
        if (logBase == 1)
            return FormulaValue.Error(ExcelError.DivisionByZero);
        if (x <= 0 || logBase <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return FormulaValue.Number(Math.Log(x, logBase));
    }

    private static FormulaValue Mod(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double y, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (y == 0)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        // Excel's MOD takes the sign of the divisor, unlike the .NET remainder operator.
        double remainder = x - (y * Math.Floor(x / y));
        return FormulaValue.Number(remainder);
    }

    private static FormulaValue Quotient(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double y, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return y == 0 ? FormulaValue.Error(ExcelError.DivisionByZero) : FormulaValue.Number(Math.Truncate(x / y));
    }

    private static FormulaValue RoundLike(EvalContext context, ReadOnlySpan<FormulaValue> args, Func<double, int, double> round)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        int digits = 0;
        if (args.Length > 1 && !Args.TryInteger(context, args[1], out digits, out error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(round(number, digits));
    }

    private static FormulaValue Step(EvalContext context, ReadOnlySpan<FormulaValue> args, bool ceiling)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double significance, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (significance == 0)
            return FormulaValue.Number(0);

        // Excel refuses a positive number with a negative step (MS-OI29500 2.1.895 and
        // 2.1.954 both understate this as "different signs"; Excel's own documented examples
        // have CEILING(-2.5, 2) = -2). A negative number with a positive step computes: the
        // ratio arithmetic lands exactly where the documentation says — CEILING pulls toward
        // zero, FLOOR away, and a negative step mirrors both.
        if (number > 0 && significance < 0)
            return FormulaValue.Error(ExcelError.Number);

        double ratio = number / significance;
        return FormulaValue.Number((ceiling ? Math.Ceiling(ratio) : Math.Floor(ratio)) * significance);
    }

    private static FormulaValue RandBetween(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int low, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int high, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return low > high
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(Random.Shared.Next(low, high + 1));
    }

    private static FormulaValue Atan2(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double x, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double y, out error))
        {
            return FormulaValue.Scalar(error);
        }

        return x == 0 && y == 0 ? FormulaValue.Error(ExcelError.DivisionByZero) : FormulaValue.Number(Math.Atan2(y, x));
    }

    private static FormulaValue Factorial(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int n, out CellValue error))
            return FormulaValue.Scalar(error);
        if (n is < 0 or > 170)
            return FormulaValue.Error(ExcelError.Number);

        double result = 1;
        for (int i = 2; i <= n; i++)
            result *= i;
        return FormulaValue.Number(result);
    }

    private static FormulaValue IntegerFold(EvalContext context, ReadOnlySpan<FormulaValue> args, Func<long, long, long> fold)
    {
        long accumulator = 0;
        bool seeded = false;
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

                if (number < 0)
                    return FormulaValue.Error(ExcelError.Number);

                long operand = (long)Math.Truncate(number);
                accumulator = seeded ? fold(accumulator, operand) : operand;
                seeded = true;
            }
        }

        return FormulaValue.Number(accumulator);
    }

    private static double Away(double value, int digits, bool up)
    {
        double factor = Math.Pow(10, digits);
        double scaled = value * factor;
        double rounded = up
            ? (scaled < 0 ? Math.Floor(scaled) : Math.Ceiling(scaled))
            : Math.Truncate(scaled);
        return rounded / factor;
    }

    private static double Parity(double value, bool even)
    {
        double sign = value < 0 ? -1 : 1;
        double magnitude = Math.Abs(value);
        double rounded = even
            ? Math.Ceiling(magnitude / 2) * 2
            : (Math.Ceiling((magnitude + 1) / 2) * 2) - 1;
        return sign * (magnitude == 0 && !even ? 1 : rounded);
    }

    private static long Gcd(long a, long b)
    {
        while (b != 0)
            (a, b) = (b, a % b);
        return a;
    }

    private static long Lcm(long a, long b) => a == 0 || b == 0 ? 0 : Math.Abs(a / Gcd(a, b) * b);
}
