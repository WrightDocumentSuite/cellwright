using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Implements the formula operators. Arithmetic on ranges and arrays broadcasts element by
/// element, which is what makes patterns like <c>SUMPRODUCT((A1:A9&gt;3)*(B1:B9))</c> work.
/// </summary>
internal static class Operators
{
    /// <summary>Applies a binary operator to two evaluated operands.</summary>
    public static FormulaValue Binary(OpCode code, in FormulaValue left, in FormulaValue right, EvalContext context)
    {
        if (code is OpCode.Range or OpCode.Union or OpCode.Intersect)
            return ReferenceOperator(code, left, right);

        if (IsMultiValued(left) || IsMultiValued(right))
            return Broadcast(code, left, right, context);

        CellValue result = Scalar(code, Dereference(left, context), Dereference(right, context), context);
        return FormulaValue.Scalar(result);
    }

    /// <summary>Applies a prefix or postfix unary operator.</summary>
    public static FormulaValue Unary(OpCode code, in FormulaValue operand, EvalContext context)
    {
        if (IsMultiValued(operand))
            return BroadcastUnary(code, operand, context);

        CellValue value = Dereference(operand, context);
        if (value.IsError)
            return FormulaValue.Scalar(value);

        CellValue number = Coercion.ToNumber(value, context.DateSystem);
        if (number.IsError)
            return FormulaValue.Scalar(number);

        double result = code switch
        {
            OpCode.Negate => -number.GetNumber(),
            OpCode.Percent => number.GetNumber() / 100,
            _ => number.GetNumber(),
        };

        return FormulaValue.Number(result);
    }

    /// <summary>Reduces a value to a single cell value, applying implicit intersection.</summary>
    public static CellValue Dereference(in FormulaValue value, EvalContext context) => value.Kind switch
    {
        FormulaValueKind.Scalar => value.AsScalar,
        FormulaValueKind.Reference => RangeReader.Intersect(value.AsReference, context),
        _ => value.AsArray.Count > 0 ? value.AsArray.Values[0] : CellValue.Error(ExcelError.Value),
    };

    private static CellValue Scalar(OpCode code, in CellValue left, in CellValue right, EvalContext context)
    {
        if (left.IsError)
            return left;
        if (right.IsError)
            return right;

        if (code == OpCode.Concat)
            return Concat(left, right, context);

        if (code >= OpCode.Equal && code <= OpCode.GreaterOrEqual)
            return Compare(code, left, right, context);

        return Arithmetic(code, left, right, context);
    }

    private static CellValue Concat(in CellValue left, in CellValue right, EvalContext context)
    {
        string text = Coercion.ToText(left, context.DateSystem) + Coercion.ToText(right, context.DateSystem);
        return text.Length > SheetLimits.MaxCellTextLength ? CellValue.Error(ExcelError.Value) : CellValue.Text(text);
    }

    private static CellValue Compare(OpCode code, in CellValue left, in CellValue right, EvalContext context)
    {
        int order = Coercion.Compare(left, right, context.DateSystem);
        bool result = code switch
        {
            OpCode.Equal => order == 0,
            OpCode.NotEqual => order != 0,
            OpCode.Less => order < 0,
            OpCode.LessOrEqual => order <= 0,
            OpCode.Greater => order > 0,
            _ => order >= 0,
        };

        return CellValue.Boolean(result);
    }

    private static CellValue Arithmetic(OpCode code, in CellValue left, in CellValue right, EvalContext context)
    {
        CellValue a = Coercion.ToNumber(left, context.DateSystem);
        if (a.IsError)
            return a;
        CellValue b = Coercion.ToNumber(right, context.DateSystem);
        if (b.IsError)
            return b;

        double x = a.GetNumber();
        double y = b.GetNumber();
        switch (code)
        {
            case OpCode.Add:
                return CellValue.Number(x + y);
            case OpCode.Subtract:
                return CellValue.Number(x - y);
            case OpCode.Multiply:
                return CellValue.Number(x * y);
            case OpCode.Divide:
                return y == 0 ? CellValue.Error(ExcelError.DivisionByZero) : CellValue.Number(x / y);
            case OpCode.Power:
                double power = Math.Pow(x, y);
                return double.IsNaN(power) ? CellValue.Error(ExcelError.Number) : CellValue.Number(power);
            default:
                return CellValue.Error(ExcelError.Value);
        }
    }

    private static FormulaValue ReferenceOperator(OpCode code, in FormulaValue left, in FormulaValue right)
    {
        // An end of the range that is itself an error carries that error out, as it does through
        // every other operator: a range whose corner was deleted is #REF!, not a type mismatch.
        if (left.Kind == FormulaValueKind.Scalar && left.AsScalar.IsError)
            return left;
        if (right.Kind == FormulaValueKind.Scalar && right.AsScalar.IsError)
            return right;

        if (left.Kind != FormulaValueKind.Reference || right.Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.Value);

        ReferenceValue a = left.AsReference;
        ReferenceValue b = right.AsReference;
        switch (code)
        {
            case OpCode.Union:
                return FormulaValue.Reference(ReferenceValue.Union(a, b));

            case OpCode.Intersect:
                ReferenceValue? intersection = ReferenceValue.Intersect(a, b);
                return intersection is null
                    ? FormulaValue.Error(ExcelError.Null)
                    : FormulaValue.Reference(intersection);

            default:
                if (!ReferenceEquals(a.Sheet, b.Sheet))
                    return FormulaValue.Error(ExcelError.Value);
                return FormulaValue.Reference(new ReferenceValue(
                    a.Sheet,
                    Math.Min(a.FirstRow, b.FirstRow),
                    Math.Min(a.FirstColumn, b.FirstColumn),
                    Math.Max(a.LastRow, b.LastRow),
                    Math.Max(a.LastColumn, b.LastColumn)));
        }
    }

    private static bool IsMultiValued(in FormulaValue value) => value.Kind switch
    {
        FormulaValueKind.Array => value.AsArray.Count > 1,
        FormulaValueKind.Reference => !value.AsReference.IsSingleCell,
        _ => false,
    };

    private static FormulaValue Broadcast(OpCode code, in FormulaValue left, in FormulaValue right, EvalContext context)
    {
        if (!TryGetArray(left, context, out ArrayValue? a) || !TryGetArray(right, context, out ArrayValue? b))
            return FormulaValue.Error(ExcelError.Value);

        int rows = Math.Max(a.Rows, b.Rows);
        int columns = Math.Max(a.Columns, b.Columns);
        if ((long)rows * columns > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Value);

        var values = new CellValue[rows * columns];
        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
                values[(row * columns) + column] = Scalar(code, a.Broadcast(row, column), b.Broadcast(row, column), context);
        }

        return FormulaValue.Array(new ArrayValue(rows, columns, values));
    }

    private static FormulaValue BroadcastUnary(OpCode code, in FormulaValue operand, EvalContext context)
    {
        if (!TryGetArray(operand, context, out ArrayValue? source))
            return FormulaValue.Error(ExcelError.Value);

        var values = new CellValue[source.Count];
        for (int i = 0; i < values.Length; i++)
            values[i] = Unary(code, FormulaValue.Scalar(source.Values[i]), context).AsScalar;

        return FormulaValue.Array(new ArrayValue(source.Rows, source.Columns, values));
    }

    /// <summary>Materializes any value as an array for broadcasting.</summary>
    public static bool TryGetArray(in FormulaValue value, EvalContext context, out ArrayValue array) =>
        TryGetArray(value, context, out array, out _);

    /// <summary>
    /// Materializes any value as an array and, for references, hands back the cache entry
    /// behind it so lookups can reuse its search index.
    /// </summary>
    public static bool TryGetArray(in FormulaValue value, EvalContext context, out ArrayValue array, out CachedRange? source)
    {
        source = null;
        switch (value.Kind)
        {
            case FormulaValueKind.Array:
                array = value.AsArray;
                return true;

            case FormulaValueKind.Reference:
                CachedRange? read = RangeReader.Materialize(value.AsReference, context);
                source = read;
                array = read?.Array ?? ArrayValue.Single(CellValue.Error(ExcelError.Value));
                return read is not null;

            default:
                array = ArrayValue.Single(value.AsScalar);
                return true;
        }
    }
}
