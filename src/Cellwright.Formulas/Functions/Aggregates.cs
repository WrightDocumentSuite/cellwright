using System.Buffers;
using System.Numerics.Tensors;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// Shared numeric aggregation. Excel treats a value written straight into a formula
/// differently from the same value sitting in a range: <c>SUM(TRUE)</c> is 1, while a TRUE
/// inside a range is skipped. Every aggregate here follows that rule.
/// </summary>
internal static class Aggregates
{
    /// <summary>Number of values worth batching into the vectorized path.</summary>
    private const int VectorThreshold = 512;

    /// <summary>
    /// Converts one aggregated value. Returns false when the value should be skipped;
    /// <paramref name="error"/> is set only when the whole aggregate must fail.
    /// </summary>
    public static bool TryNumeric(
        in CellValue value, bool fromRange, EvalContext context, out double number, out CellValue error)
    {
        number = 0;
        error = CellValue.Blank;
        switch (value.Kind)
        {
            case CellValueKind.Number:
                number = value.GetNumber();
                return true;

            case CellValueKind.DateTime or CellValueKind.TimeSpan:
                number = value.ToSerialNumber(context.DateSystem);
                return true;

            case CellValueKind.Error:
                error = value;
                return false;

            case CellValueKind.Blank:
                return false;

            case CellValueKind.Boolean:
                if (fromRange)
                    return false;
                number = value.GetBoolean() ? 1 : 0;
                return true;

            default:
                if (fromRange)
                    return false;
                CellValue parsed = Coercion.ToNumber(value, context.DateSystem);
                if (parsed.IsError)
                {
                    error = parsed;
                    return false;
                }

                number = parsed.GetNumber();
                return true;
        }
    }

    /// <summary>Running statistics over the numeric values of a set of arguments.</summary>
    public struct Accumulator
    {
        public double Sum;
        public double Min = double.MaxValue;
        public double Max = double.MinValue;
        public int Count;
        public int NonEmptyCount;
        public CellValue Error = CellValue.Blank;

        public Accumulator()
        {
        }

        public readonly bool Failed => Error.IsError;

        public void Add(double value)
        {
            Sum += value;
            Min = Math.Min(Min, value);
            Max = Math.Max(Max, value);
            Count++;
        }
    }

    /// <summary>Accumulates every numeric value found in the arguments.</summary>
    public static Accumulator Collect(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        var accumulator = new Accumulator();
        foreach (ref readonly FormulaValue arg in args)
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!value.IsBlank)
                    accumulator.NonEmptyCount++;

                if (TryNumeric(value, fromRange, context, out double number, out CellValue error))
                {
                    accumulator.Add(number);
                }
                else if (error.IsError)
                {
                    accumulator.Error = error;
                    return accumulator;
                }
            }
        }

        return accumulator;
    }

    /// <summary>
    /// Sums the numbers of a single reference. Large ranges are copied into a pooled buffer
    /// and reduced with hardware vectors; smaller ones take the plain loop, which avoids the
    /// copy entirely.
    /// </summary>
    public static double SumReference(ReferenceValue reference, EvalContext context)
    {
        long cells = RangeReader.CountCells(reference);
        if (cells < VectorThreshold)
            return SumScalar(reference, context);

        double[] buffer = ArrayPool<double>.Shared.Rent((int)Math.Min(cells, VectorThreshold * 8));
        try
        {
            double total = 0;
            int count = 0;
            foreach (CellValue value in new ReferenceEnumerator(reference, context))
            {
                if (!value.IsNumber)
                {
                    if (value.IsError)
                        return double.NaN;
                    continue;
                }

                buffer[count++] = value.GetNumber();
                if (count == buffer.Length)
                {
                    total += TensorPrimitives.Sum<double>(buffer);
                    count = 0;
                }
            }

            return total + TensorPrimitives.Sum<double>(buffer.AsSpan(0, count));
        }
        finally
        {
            ArrayPool<double>.Shared.Return(buffer);
        }
    }

    private static double SumScalar(ReferenceValue reference, EvalContext context)
    {
        double total = 0;
        foreach (CellValue value in new ReferenceEnumerator(reference, context))
        {
            if (value.IsNumber)
                total += value.GetNumber();
            else if (value.IsError)
                return double.NaN;
        }

        return total;
    }
}
