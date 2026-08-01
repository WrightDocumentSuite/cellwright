namespace Cellwright.Formulas.Functions;

/// <summary>
/// The engineering catalog of §18.17.7: number bases in two's complement, the step functions,
/// and <c>CONVERT</c> with its unit table.
/// </summary>
internal static class EngineeringFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("BIN2DEC", 1, 1, (c, a) => FromBase(c, a, 2));
        builder.Add("OCT2DEC", 1, 1, (c, a) => FromBase(c, a, 8));
        builder.Add("HEX2DEC", 1, 1, (c, a) => FromBase(c, a, 16));
        builder.Add("DEC2BIN", 1, 2, (c, a) => ToBase(c, a, 2));
        builder.Add("DEC2OCT", 1, 2, (c, a) => ToBase(c, a, 8));
        builder.Add("DEC2HEX", 1, 2, (c, a) => ToBase(c, a, 16));
        builder.Add("BIN2OCT", 1, 2, (c, a) => Rebase(c, a, 2, 8));
        builder.Add("BIN2HEX", 1, 2, (c, a) => Rebase(c, a, 2, 16));
        builder.Add("OCT2BIN", 1, 2, (c, a) => Rebase(c, a, 8, 2));
        builder.Add("OCT2HEX", 1, 2, (c, a) => Rebase(c, a, 8, 16));
        builder.Add("HEX2BIN", 1, 2, (c, a) => Rebase(c, a, 16, 2));
        builder.Add("HEX2OCT", 1, 2, (c, a) => Rebase(c, a, 16, 8));
        builder.Add("DELTA", 1, 2, Delta);
        builder.Add("GESTEP", 1, 2, GeStep);
        builder.Add("CONVERT", 3, 3, UnitConversions.Convert);
    }

    /// <summary>Bits of the fixed two's-complement window each base works in.</summary>
    private static int Bits(int radix) => radix switch
    {
        2 => 10,   // ten binary digits
        8 => 30,   // ten octal digits
        _ => 40,   // ten hexadecimal digits
    };

    private static FormulaValue FromBase(EvalContext context, ReadOnlySpan<FormulaValue> args, int radix)
    {
        if (!TryParse(context, args[0], radix, out long value, out CellValue error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(value);
    }

    private static FormulaValue ToBase(EvalContext context, ReadOnlySpan<FormulaValue> args, int radix)
    {
        if (!Args.TryNumber(context, args[0], out double raw, out CellValue error))
            return FormulaValue.Scalar(error);

        long value = (long)Math.Truncate(raw);
        return Render(context, args, value, radix);
    }

    private static FormulaValue Rebase(EvalContext context, ReadOnlySpan<FormulaValue> args, int from, int to)
    {
        if (!TryParse(context, args[0], from, out long value, out CellValue error))
            return FormulaValue.Scalar(error);

        return Render(context, args, value, to);
    }

    /// <summary>Reads a based number of up to ten digits, the top bit carrying the sign.</summary>
    private static bool TryParse(EvalContext context, in FormulaValue argument, int radix, out long value, out CellValue error)
    {
        value = 0;
        string text = Args.Text(context, argument).Trim();
        error = CellValue.Error(ExcelError.Number);
        if (text.Length is 0 or > 10)
            return false;

        long magnitude = 0;
        foreach (char c in text)
        {
            int digit = char.ToUpperInvariant(c) switch
            {
                >= '0' and <= '9' => c - '0',
                >= 'A' and <= 'F' => char.ToUpperInvariant(c) - 'A' + 10,
                _ => -1,
            };
            if (digit < 0 || digit >= radix)
                return false;

            magnitude = (magnitude * radix) + digit;
        }

        // Ten digits span a fixed number of bits, and the highest bit is the sign.
        long window = 1L << Bits(radix);
        value = magnitude >= window / 2 ? magnitude - window : magnitude;
        error = CellValue.Blank;
        return true;
    }

    /// <summary>Writes a number in the base's ten-digit window, padded when places asks.</summary>
    private static FormulaValue Render(EvalContext context, ReadOnlySpan<FormulaValue> args, long value, int radix)
    {
        long window = 1L << Bits(radix);
        if (value < -(window / 2) || value >= window / 2)
            return FormulaValue.Error(ExcelError.Number);

        // A negative number fills the whole window: places is ignored, as in Excel.
        string digits = ToDigits(value < 0 ? value + window : value, radix);
        if (value < 0)
            return FormulaValue.Text(digits.PadLeft(10, '0'));

        if (args.Length > 1 && !args[1].IsBlank)
        {
            if (!Args.TryInteger(context, args[1], out int places, out CellValue error))
                return FormulaValue.Scalar(error);
            if (places < 0)
                return FormulaValue.Error(ExcelError.Number);
            if (digits.Length > places || places > 10)
                return FormulaValue.Error(ExcelError.Number);

            digits = digits.PadLeft(places, '0');
        }

        return FormulaValue.Text(digits);
    }

    private static string ToDigits(long value, int radix)
    {
        if (value == 0)
            return "0";

        Span<char> buffer = stackalloc char[41];
        int at = buffer.Length;
        while (value > 0)
        {
            long digit = value % radix;
            buffer[--at] = (char)(digit < 10 ? '0' + digit : 'A' + digit - 10);
            value /= radix;
        }

        return new string(buffer[at..]);
    }

    private static FormulaValue Delta(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double first, out CellValue error))
            return FormulaValue.Scalar(error);

        double second = 0;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryNumber(context, args[1], out second, out error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(first == second ? 1 : 0);
    }

    private static FormulaValue GeStep(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        double step = 0;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryNumber(context, args[1], out step, out error))
            return FormulaValue.Scalar(error);

        return FormulaValue.Number(number >= step ? 1 : 0);
    }
}

/// <summary>
/// The unit table of <c>CONVERT</c>. Every unit reduces to one factor against its category's
/// base unit; temperatures are the exception, converted through Kelvin with their offsets.
/// Metric units accept the decimal prefixes, bits and bytes also the binary ones.
/// </summary>
internal static class UnitConversions
{
    /// <summary>A unit: its category, its size in the category's base unit, prefix rules.</summary>
    private readonly record struct Unit(string Category, double Factor, bool Metric, bool Binary = false);

    private static readonly Dictionary<string, Unit> Units = BuildUnits();

    private static readonly (string Prefix, double Factor)[] DecimalPrefixes =
    [
        ("Y", 1e24), ("Z", 1e21), ("E", 1e18), ("P", 1e15), ("T", 1e12), ("G", 1e9), ("M", 1e6),
        ("k", 1e3), ("h", 1e2), ("da", 1e1), ("e", 1e1), ("d", 1e-1), ("c", 1e-2), ("m", 1e-3),
        ("u", 1e-6), ("n", 1e-9), ("p", 1e-12), ("f", 1e-15), ("a", 1e-18), ("z", 1e-21), ("y", 1e-24),
    ];

    private static readonly (string Prefix, double Factor)[] BinaryPrefixes =
    [
        ("Yi", Math.Pow(2, 80)), ("Zi", Math.Pow(2, 70)), ("Ei", Math.Pow(2, 60)), ("Pi", Math.Pow(2, 50)),
        ("Ti", Math.Pow(2, 40)), ("Gi", Math.Pow(2, 30)), ("Mi", Math.Pow(2, 20)), ("ki", Math.Pow(2, 10)),
    ];

    public static FormulaValue Convert(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        string fromText = Args.Text(context, args[1]);
        string toText = Args.Text(context, args[2]);
        if (!TryResolve(fromText, out Unit from, out double fromScale) ||
            !TryResolve(toText, out Unit to, out double toScale))
        {
            return FormulaValue.Error(ExcelError.NotAvailable);
        }

        if (from.Category != to.Category)
            return FormulaValue.Error(ExcelError.NotAvailable);

        if (from.Category == "temperature")
            return FormulaValue.Number(FromKelvin(ToKelvin(number, fromText), toText));

        double result = number * from.Factor * fromScale / (to.Factor * toScale);
        return double.IsNaN(result) || double.IsInfinity(result)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(result);
    }

    /// <summary>Splits an optional prefix off a unit name. Names are case-sensitive, as in Excel.</summary>
    private static bool TryResolve(string text, out Unit unit, out double scale)
    {
        scale = 1;
        if (Units.TryGetValue(text, out unit))
            return true;

        foreach ((string prefix, double factor) in BinaryPrefixes)
        {
            if (text.StartsWith(prefix, StringComparison.Ordinal) &&
                Units.TryGetValue(text[prefix.Length..], out unit) && unit.Binary)
            {
                scale = factor;
                return true;
            }
        }

        foreach ((string prefix, double factor) in DecimalPrefixes)
        {
            if (text.StartsWith(prefix, StringComparison.Ordinal) &&
                Units.TryGetValue(text[prefix.Length..], out unit) && unit.Metric)
            {
                scale = factor;
                return true;
            }
        }

        return false;
    }

    private static double ToKelvin(double value, string unit) => unit switch
    {
        "C" or "cel" => value + 273.15,
        "F" or "fah" => (value + 459.67) * 5 / 9,
        "Rank" => value * 5 / 9,
        "Reau" => (value * 1.25) + 273.15,
        _ => value,
    };

    private static double FromKelvin(double kelvin, string unit) => unit switch
    {
        "C" or "cel" => kelvin - 273.15,
        "F" or "fah" => (kelvin * 9 / 5) - 459.67,
        "Rank" => kelvin * 9 / 5,
        "Reau" => (kelvin - 273.15) * 0.8,
        _ => kelvin,
    };

    private static Dictionary<string, Unit> BuildUnits()
    {
        var units = new Dictionary<string, Unit>(StringComparer.Ordinal);
        void Add(string category, double factor, bool metric, params string[] names)
        {
            foreach (string name in names)
                units[name] = new Unit(category, factor, metric);
        }

        // Weight, in grams.
        Add("weight", 1, true, "g");
        Add("weight", 14593.90293720636, false, "sg");
        Add("weight", 453.59237, false, "lbm");
        Add("weight", 1.660538782e-24, true, "u");
        Add("weight", 28.349523125, false, "ozm");
        Add("weight", 0.06479891, false, "grain");
        Add("weight", 45359.237, false, "cwt", "shweight");
        Add("weight", 50802.34544, false, "uk_cwt", "lcwt", "hweight");
        Add("weight", 6350.29318, false, "stone");
        Add("weight", 907184.74, false, "ton");
        Add("weight", 1016046.9088, false, "uk_ton", "LTON", "brton");

        // Distance, in meters.
        Add("distance", 1, true, "m");
        Add("distance", 1609.344, false, "mi");
        Add("distance", 1852, false, "Nmi");
        Add("distance", 0.0254, false, "in");
        Add("distance", 0.3048, false, "ft");
        Add("distance", 0.9144, false, "yd");
        Add("distance", 1e-10, true, "ang");
        Add("distance", 1.143, false, "ell");
        Add("distance", 9.46073047258e15, true, "ly");
        Add("distance", 3.08567758128e16, true, "parsec", "pc");
        Add("distance", 0.0254 / 72, false, "Picapt", "Pica");
        Add("distance", 0.0254 / 6, false, "pica");
        Add("distance", 1609.3472186944375, false, "survey_mi");

        // Time, in seconds.
        Add("time", 31557600, false, "yr");
        Add("time", 86400, false, "day", "d");
        Add("time", 3600, false, "hr");
        Add("time", 60, false, "mn", "min");
        Add("time", 1, true, "sec", "s");

        // Pressure, in pascals.
        Add("pressure", 1, true, "Pa", "p");
        Add("pressure", 101325, true, "atm", "at");
        Add("pressure", 133.322, true, "mmHg");
        Add("pressure", 6894.757293168361, false, "psi");
        Add("pressure", 133.32236842105263, false, "Torr");

        // Force, in newtons.
        Add("force", 1, true, "N");
        Add("force", 1e-5, true, "dyn", "dy");
        Add("force", 4.4482216152605, false, "lbf");
        Add("force", 0.00980665, false, "pond");

        // Energy, in joules.
        Add("energy", 1, true, "J");
        Add("energy", 1e-7, true, "e");
        Add("energy", 4.184, true, "c");
        Add("energy", 4.1868, true, "cal");
        Add("energy", 1.602176487e-19, true, "eV", "ev");
        Add("energy", 2684519.5368856, false, "HPh", "hh");
        Add("energy", 3600, true, "Wh", "wh");
        Add("energy", 1.3558179483314004, false, "flb");
        Add("energy", 1055.05585262, false, "BTU", "btu");

        // Power, in watts.
        Add("power", 745.69987158227022, false, "HP", "h");
        Add("power", 735.49875, false, "PS");
        Add("power", 1, true, "W", "w");

        // Magnetism, in teslas.
        Add("magnetism", 1, true, "T");
        Add("magnetism", 1e-4, true, "ga");

        // Temperature converts through Kelvin; the factor only marks the category.
        Add("temperature", 1, false, "C", "cel", "F", "fah", "K", "kel", "Rank", "Reau");
        units["K"] = units["K"] with { Metric = true };
        units["kel"] = units["kel"] with { Metric = true };

        // Liquid measure, in liters.
        Add("liquid", 0.00492892159375, false, "tsp");
        Add("liquid", 0.005, false, "tspm");
        Add("liquid", 0.01478676478125, false, "tbs");
        Add("liquid", 0.0295735295625, false, "oz");
        Add("liquid", 0.2365882365, false, "cup");
        Add("liquid", 0.473176473, false, "pt", "us_pt");
        Add("liquid", 0.56826125, false, "uk_pt");
        Add("liquid", 0.946352946, false, "qt");
        Add("liquid", 1.1365225, false, "uk_qt");
        Add("liquid", 3.785411784, false, "gal");
        Add("liquid", 4.54609, false, "uk_gal");
        Add("liquid", 1, true, "l", "L", "lt");

        // Information, where the binary prefixes apply.
        units["bit"] = new Unit("information", 1, Metric: true, Binary: true);
        units["byte"] = new Unit("information", 8, Metric: true, Binary: true);

        // Speed, in meters per second.
        Add("speed", 1, true, "m/s", "m/sec");
        Add("speed", 1.0 / 3600, true, "m/h", "m/hr");
        Add("speed", 0.44704, false, "mph");
        Add("speed", 1852.0 / 3600, false, "kn");
        Add("speed", 1853.184 / 3600, false, "admkn");

        return units;
    }
}
