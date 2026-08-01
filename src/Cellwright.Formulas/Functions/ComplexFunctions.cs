using System.Globalization;
using System.Numerics;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// Complex numbers the way the format sees them: pieces of text like <c>3+4i</c>. Every
/// function parses its arguments, computes on <see cref="Complex"/>, and renders back with
/// the suffix the input used — <c>i</c> and <c>j</c> never mix inside one call.
/// </summary>
internal static class ComplexFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("COMPLEX", 2, 3, Construct);
        builder.Add("IMREAL", 1, 1, (c, a) => Part(c, a, static z => z.Real));
        builder.Add("IMAGINARY", 1, 1, (c, a) => Part(c, a, static z => z.Imaginary));
        builder.Add("IMABS", 1, 1, (c, a) => Part(c, a, static z => z.Magnitude));
        builder.Add("IMARGUMENT", 1, 1, Argument);
        builder.Add("IMCONJUGATE", 1, 1, (c, a) => Map(c, a, Complex.Conjugate));
        builder.Add("IMSUM", 1, FunctionDefinition.Unlimited, (c, a) => Fold(c, a, static (x, y) => x + y));
        builder.Add("IMPRODUCT", 1, FunctionDefinition.Unlimited, (c, a) => Fold(c, a, static (x, y) => x * y));
        builder.Add("IMSUB", 2, 2, (c, a) => Combine(c, a, static (x, y) => x - y));
        builder.Add("IMDIV", 2, 2, Divide);
        builder.Add("IMPOWER", 2, 2, Power);
        builder.Add("IMSQRT", 1, 1, (c, a) => Map(c, a, Complex.Sqrt));
        builder.Add("IMEXP", 1, 1, (c, a) => Map(c, a, Complex.Exp));
        builder.Add("IMLN", 1, 1, (c, a) => MapNonZero(c, a, Complex.Log));
        builder.Add("IMLOG10", 1, 1, (c, a) => MapNonZero(c, a, static z => Complex.Log(z) / Math.Log(10)));
        builder.Add("IMLOG2", 1, 1, (c, a) => MapNonZero(c, a, static z => Complex.Log(z) / Math.Log(2)));
        builder.Add("IMSIN", 1, 1, (c, a) => Map(c, a, Complex.Sin));
        builder.Add("IMCOS", 1, 1, (c, a) => Map(c, a, Complex.Cos));
        builder.Add("IMTAN", 1, 1, (c, a) => Map(c, a, Complex.Tan));
    }

    private static FormulaValue Construct(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double real, out CellValue error) ||
            !Args.TryNumber(context, args[1], out double imaginary, out error))
        {
            return FormulaValue.Scalar(error);
        }

        char suffix = 'i';
        if (args.Length > 2 && !args[2].IsBlank)
        {
            string text = Args.Text(context, args[2]);
            if (text is not ("i" or "j"))
                return FormulaValue.Error(ExcelError.Value);
            suffix = text[0];
        }

        return FormulaValue.Text(Render(new Complex(real, imaginary), suffix));
    }

    private static FormulaValue Part(EvalContext context, ReadOnlySpan<FormulaValue> args, Func<Complex, double> part)
    {
        if (!TryParse(context, args[0], out Complex z, out _, out CellValue error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(part(z));
    }

    private static FormulaValue Argument(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryParse(context, args[0], out Complex z, out _, out CellValue error))
            return FormulaValue.Scalar(error);
        if (z == Complex.Zero)
            return FormulaValue.Error(ExcelError.DivisionByZero);

        return FormulaValue.Number(z.Phase);
    }

    private static FormulaValue Map(EvalContext context, ReadOnlySpan<FormulaValue> args, Func<Complex, Complex> map)
    {
        if (!TryParse(context, args[0], out Complex z, out char suffix, out CellValue error))
            return FormulaValue.Scalar(error);

        return Wrap(map(z), suffix);
    }

    private static FormulaValue MapNonZero(EvalContext context, ReadOnlySpan<FormulaValue> args, Func<Complex, Complex> map)
    {
        if (!TryParse(context, args[0], out Complex z, out char suffix, out CellValue error))
            return FormulaValue.Scalar(error);
        if (z == Complex.Zero)
            return FormulaValue.Error(ExcelError.Number);

        return Wrap(map(z), suffix);
    }

    private static FormulaValue Fold(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Func<Complex, Complex, Complex> fold)
    {
        Complex? accumulated = null;
        char suffix = '\0';
        foreach (ref readonly FormulaValue arg in args)
        {
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsError)
                    return FormulaValue.Scalar(value);
                if (value.IsBlank)
                    continue;
                if (!TryParseValue(value, out Complex z, out char seen))
                    return FormulaValue.Error(ExcelError.Number);
                if (!Blend(ref suffix, seen))
                    return FormulaValue.Error(ExcelError.Value);

                accumulated = accumulated is { } sum ? fold(sum, z) : z;
            }
        }

        return accumulated is { } result ? Wrap(result, suffix) : FormulaValue.Error(ExcelError.Value);
    }

    private static FormulaValue Combine(
        EvalContext context, ReadOnlySpan<FormulaValue> args, Func<Complex, Complex, Complex> combine)
    {
        if (!TryParse(context, args[0], out Complex x, out char first, out CellValue error) ||
            !TryParse(context, args[1], out Complex y, out char second, out error))
        {
            return FormulaValue.Scalar(error);
        }

        char suffix = '\0';
        if (!Blend(ref suffix, first) || !Blend(ref suffix, second))
            return FormulaValue.Error(ExcelError.Value);

        return Wrap(combine(x, y), suffix);
    }

    private static FormulaValue Divide(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryParse(context, args[0], out Complex x, out char first, out CellValue error) ||
            !TryParse(context, args[1], out Complex y, out char second, out error))
        {
            return FormulaValue.Scalar(error);
        }

        if (y == Complex.Zero)
            return FormulaValue.Error(ExcelError.Number);

        char suffix = '\0';
        if (!Blend(ref suffix, first) || !Blend(ref suffix, second))
            return FormulaValue.Error(ExcelError.Value);

        return Wrap(x / y, suffix);
    }

    private static FormulaValue Power(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryParse(context, args[0], out Complex z, out char suffix, out CellValue error))
            return FormulaValue.Scalar(error);
        if (!Args.TryNumber(context, args[1], out double exponent, out CellValue bad))
            return FormulaValue.Scalar(bad);
        if (z == Complex.Zero && exponent <= 0)
            return FormulaValue.Error(ExcelError.Number);

        return Wrap(Complex.Pow(z, exponent), suffix);
    }

    /// <summary>
    /// Two suffixes agree when equal or when one side never named one — a bare real number
    /// takes whichever letter its neighbours use. Naming both letters in one call refuses.
    /// </summary>
    private static bool Blend(ref char suffix, char seen)
    {
        if (seen == '\0')
            return true;
        if (suffix == '\0')
        {
            suffix = seen;
            return true;
        }

        return suffix == seen;
    }

    private static FormulaValue Wrap(Complex z, char suffix)
    {
        if (double.IsNaN(z.Real) || double.IsInfinity(z.Real) ||
            double.IsNaN(z.Imaginary) || double.IsInfinity(z.Imaginary))
        {
            return FormulaValue.Error(ExcelError.Number);
        }

        return FormulaValue.Text(Render(z, suffix == '\0' ? 'i' : suffix));
    }

    private static bool TryParse(
        EvalContext context, in FormulaValue argument, out Complex value, out char suffix, out CellValue error)
    {
        CellValue scalar = Operators.Dereference(argument, context);
        if (scalar.IsError)
        {
            value = default;
            suffix = '\0';
            error = scalar;
            return false;
        }

        if (TryParseValue(scalar, out value, out suffix))
        {
            error = CellValue.Blank;
            return true;
        }

        error = CellValue.Error(ExcelError.Number);
        return false;
    }

    /// <summary>Reads a complex value from a cell: a plain number, or text like <c>-3.5+2e2j</c>.</summary>
    private static bool TryParseValue(in CellValue scalar, out Complex value, out char suffix)
    {
        value = default;
        suffix = '\0';
        if (scalar.Kind == CellValueKind.Number)
        {
            value = new Complex(scalar.GetNumber(), 0);
            return true;
        }

        if (scalar.IsBlank)
            return true;
        if (!scalar.IsText)
            return false;

        string text = scalar.GetText().Trim();
        if (text.Length == 0)
            return false;

        char last = text[^1];
        if (last is not ('i' or 'j'))
        {
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double real))
                return false;
            value = new Complex(real, 0);
            return true;
        }

        suffix = last;
        string body = text[..^1];

        // The split between the real and the imaginary part is the last sign that neither
        // starts the text nor follows an exponent marker.
        int split = -1;
        for (int i = body.Length - 1; i > 0; i--)
        {
            if (body[i] is not ('+' or '-'))
                continue;
            if (body[i - 1] is 'e' or 'E')
                continue;
            split = i;
            break;
        }

        string realText = split < 0 ? "" : body[..split];
        string imaginaryText = split < 0 ? body : body[split..];
        double imaginary = imaginaryText switch
        {
            "" or "+" => 1,
            "-" => -1,
            _ => double.TryParse(imaginaryText, NumberStyles.Float, CultureInfo.InvariantCulture, out double part)
                ? part
                : double.NaN,
        };
        if (double.IsNaN(imaginary))
            return false;

        double realPart = 0;
        if (realText.Length > 0 &&
            !double.TryParse(realText, NumberStyles.Float, CultureInfo.InvariantCulture, out realPart))
        {
            return false;
        }

        value = new Complex(realPart, imaginary);
        return true;
    }

    /// <summary>Writes a complex value the way Excel spells one: <c>3-4i</c>, <c>i</c>, <c>0</c>.</summary>
    internal static string Render(Complex z, char suffix)
    {
        string real = Number(z.Real);
        if (z.Imaginary == 0)
            return real;

        string imaginary = z.Imaginary switch
        {
            1 => "",
            -1 => "-",
            _ => Number(z.Imaginary),
        };

        if (z.Real == 0)
            return imaginary + suffix;

        string sign = z.Imaginary > 0 && imaginary.Length > 0 && imaginary[0] != '-' ? "+" : imaginary.Length == 0 ? "+" : "";
        return real + sign + imaginary + suffix;
    }

    private static string Number(double value) => value.ToString("G15", CultureInfo.InvariantCulture);
}
