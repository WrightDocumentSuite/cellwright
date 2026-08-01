namespace Cellwright.Formulas.Functions;

/// <summary>
/// The matrix trio of §18.17.7: multiply, invert, take the determinant. All three demand every
/// cell be a number — a blank or a word anywhere in the range is <c>#VALUE!</c>, as in Excel —
/// and the first two return arrays, showing their top-left corner in a scalar context.
/// </summary>
internal static class MatrixFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("MMULT", 2, 2, Multiply, wantsRawReferences: true);
        builder.Add("MINVERSE", 1, 1, Inverse, wantsRawReferences: true);
        builder.Add("MDETERM", 1, 1, Determinant, wantsRawReferences: true);
    }

    private static FormulaValue Multiply(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (ToMatrix(context, args[0]) is not { } left || ToMatrix(context, args[1]) is not { } right)
            return FormulaValue.Error(ExcelError.Value);
        if (left.GetLength(1) != right.GetLength(0))
            return FormulaValue.Error(ExcelError.Value);

        int rows = left.GetLength(0);
        int inner = left.GetLength(1);
        int columns = right.GetLength(1);
        var values = new CellValue[rows * columns];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                double sum = 0;
                for (int k = 0; k < inner; k++)
                    sum += left[r, k] * right[k, c];
                values[(r * columns) + c] = CellValue.Number(sum);
            }
        }

        return FormulaValue.Array(new ArrayValue(rows, columns, values));
    }

    private static FormulaValue Inverse(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (ToMatrix(context, args[0]) is not { } matrix || matrix.GetLength(0) != matrix.GetLength(1))
            return FormulaValue.Error(ExcelError.Value);
        if (!Matrix.TryInvert(matrix, out double[,] inverse))
            return FormulaValue.Error(ExcelError.Number);

        int size = matrix.GetLength(0);
        var values = new CellValue[size * size];
        for (int r = 0; r < size; r++)
        {
            for (int c = 0; c < size; c++)
                values[(r * size) + c] = CellValue.Number(inverse[r, c]);
        }

        return FormulaValue.Array(new ArrayValue(size, size, values));
    }

    private static FormulaValue Determinant(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (ToMatrix(context, args[0]) is not { } matrix || matrix.GetLength(0) != matrix.GetLength(1))
            return FormulaValue.Error(ExcelError.Value);

        return FormulaValue.Number(Matrix.Determinant(matrix));
    }

    /// <summary>The argument as a dense numeric matrix, or null when any cell is not a number.</summary>
    private static double[,]? ToMatrix(EvalContext context, in FormulaValue value)
    {
        if (!Operators.TryGetArray(value, context, out ArrayValue array))
            return null;

        var matrix = new double[array.Rows, array.Columns];
        for (int r = 0; r < array.Rows; r++)
        {
            for (int c = 0; c < array.Columns; c++)
            {
                CellValue cell = array[r, c];
                if (cell.Kind is not (CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan))
                    return null;

                matrix[r, c] = cell.ToSerialNumber(context.DateSystem);
            }
        }

        return matrix;
    }
}
