namespace Cellwright.Formulas.Functions;

/// <summary>
/// Walks every scalar inside one argument, whether it is a literal, a reference or an array.
/// <see cref="FromRange"/> distinguishes the two cases that Excel treats differently: text and
/// logicals written directly in a formula are converted, while the same values inside a range
/// are skipped by numeric aggregates.
/// </summary>
internal ref struct ArgumentValues
{
    private readonly ArrayValue? _array;
    private ReferenceEnumerator _reference;
    private CellValue _scalar;
    private int _index;
    private readonly Source _source;

    public ArgumentValues(in FormulaValue value, EvalContext context)
    {
        switch (value.Kind)
        {
            case FormulaValueKind.Reference:
                _reference = new ReferenceEnumerator(value.AsReference, context);
                _source = Source.Reference;
                break;
            case FormulaValueKind.Array:
                _array = value.AsArray;
                _source = Source.Array;
                break;
            default:
                _scalar = value.AsScalar;
                _source = Source.Scalar;
                break;
        }

        _index = -1;
    }

    /// <summary>The value the enumerator stands on.</summary>
    public CellValue Current { get; private set; }

    /// <summary>Whether the current value came from a range or array rather than a literal.</summary>
    public readonly bool FromRange => _source != Source.Scalar;

    public ArgumentValues GetEnumerator() => this;

    public bool MoveNext()
    {
        switch (_source)
        {
            case Source.Scalar:
                if (_index >= 0)
                    return false;
                _index = 0;
                Current = _scalar;
                return true;

            case Source.Array:
                if (++_index >= _array!.Count)
                    return false;
                Current = _array.Values[_index];
                return true;

            default:
                if (!_reference.MoveNext())
                    return false;
                Current = _reference.Current;
                return true;
        }
    }

    private enum Source : byte
    {
        Scalar,
        Array,
        Reference,
    }
}

/// <summary>
/// Argument conversions shared by the function catalogs.
/// </summary>
internal static class Args
{
    /// <summary>Reads an argument as a number, reporting Excel's error instead of throwing.</summary>
    public static bool TryNumber(EvalContext context, in FormulaValue value, out double number, out CellValue error)
    {
        CellValue scalar = Operators.Dereference(value, context);
        CellValue converted = Coercion.ToNumber(scalar, context.DateSystem);
        if (converted.IsError)
        {
            number = 0;
            error = converted;
            return false;
        }

        number = converted.GetNumber();
        error = CellValue.Blank;
        return true;
    }

    /// <summary>Reads an argument as an integer, truncating toward zero the way Excel does.</summary>
    public static bool TryInteger(EvalContext context, in FormulaValue value, out int result, out CellValue error)
    {
        result = 0;
        if (!TryNumber(context, value, out double number, out error))
            return false;

        double truncated = Math.Truncate(number);
        if (truncated is > int.MaxValue or < int.MinValue)
        {
            error = CellValue.Error(ExcelError.Number);
            return false;
        }

        result = (int)truncated;
        return true;
    }

    /// <summary>Reads an argument as text.</summary>
    public static string Text(EvalContext context, in FormulaValue value) =>
        Coercion.ToText(Operators.Dereference(value, context), context.DateSystem);

    /// <summary>Reads an argument as a logical value.</summary>
    public static bool TryBoolean(EvalContext context, in FormulaValue value, out bool result, out CellValue error)
    {
        CellValue converted = Coercion.ToBoolean(Operators.Dereference(value, context));
        if (converted.IsError)
        {
            result = false;
            error = converted;
            return false;
        }

        result = converted.GetBoolean();
        error = CellValue.Blank;
        return true;
    }

    /// <summary>Returns the argument at <paramref name="index"/>, or blank when omitted.</summary>
    public static FormulaValue At(ReadOnlySpan<FormulaValue> args, int index) =>
        index < args.Length ? args[index] : FormulaValue.Blank;

    /// <summary>Runs a numeric function over one argument, propagating conversion errors.</summary>
    public static FormulaValue Map(EvalContext context, in FormulaValue value, Func<double, double> map)
    {
        if (!TryNumber(context, value, out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        double result = map(number);
        return double.IsNaN(result) || double.IsInfinity(result)
            ? FormulaValue.Error(ExcelError.Number)
            : FormulaValue.Number(result);
    }

    /// <summary>
    /// Excel rounds halves away from zero, unlike the banker's rounding .NET defaults to.
    /// It also rounds the fifteen-significant-digit view of the number, which is why
    /// <c>ROUND(1.005,2)</c> is 1.01 there but 1.00 with plain binary arithmetic. Converting
    /// through <see cref="decimal"/> reproduces both behaviours exactly.
    /// </summary>
    public static double Round(double value, int digits)
    {
        if (digits is >= 0 and <= 28 && Math.Abs(value) < 7.9e28)
            return (double)Math.Round((decimal)value, digits, MidpointRounding.AwayFromZero);

        double factor = Math.Pow(10, digits);
        double scaled = value * factor;
        if (double.IsInfinity(scaled))
            return value;
        return Math.Round(scaled, MidpointRounding.AwayFromZero) / factor;
    }
}
