namespace Cellwright.Formulas.Functions;

/// <summary>
/// Excel 2022–2024 dynamic array shaping, stacking, slicing, and wrapping functions:
/// <c>VSTACK</c>, <c>HSTACK</c>, <c>TOCOL</c>, <c>TOROW</c>, <c>TAKE</c>, <c>DROP</c>,
/// <c>CHOOSEROWS</c>, <c>CHOOSECOLS</c>, <c>WRAPROWS</c>, <c>WRAPCOLS</c>, <c>EXPAND</c>.
/// </summary>
internal static class ArrayShapeFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("VSTACK", 1, FunctionDefinition.Unlimited, (c, a) => Stack(c, a, vertical: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("HSTACK", 1, FunctionDefinition.Unlimited, (c, a) => Stack(c, a, vertical: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("TOCOL", 1, 3, (c, a) => Flatten(c, a, toColumn: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("TOROW", 1, 3, (c, a) => Flatten(c, a, toColumn: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("TAKE", 2, 3, (c, a) => Slice(c, a, isTake: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("DROP", 2, 3, (c, a) => Slice(c, a, isTake: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("CHOOSEROWS", 2, FunctionDefinition.Unlimited, (c, a) => Choose(c, a, byColumn: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("CHOOSECOLS", 2, FunctionDefinition.Unlimited, (c, a) => Choose(c, a, byColumn: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("WRAPROWS", 2, 3, (c, a) => Wrap(c, a, byRows: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("WRAPCOLS", 2, 3, (c, a) => Wrap(c, a, byRows: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("EXPAND", 2, 4, Expand, wantsRawReferences: true, acceptsErrors: true);
    }

    private static FormulaValue Stack(EvalContext context, ReadOnlySpan<FormulaValue> args, bool vertical)
    {
        var arrays = new List<ArrayValue>(args.Length);
        int totalPrimary = 0;
        int maxSecondary = 0;
        foreach (ref readonly FormulaValue arg in args)
        {
            if (arg.IsError)
                return arg;
            if (!Operators.TryGetArray(arg, context, out ArrayValue array))
                return FormulaValue.Error(ExcelError.Value);

            arrays.Add(array);
            totalPrimary += vertical ? array.Rows : array.Columns;
            maxSecondary = Math.Max(maxSecondary, vertical ? array.Columns : array.Rows);
        }

        int rows = vertical ? totalPrimary : maxSecondary;
        int cols = vertical ? maxSecondary : totalPrimary;
        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        return CopyStacked(arrays, rows, cols, vertical);
    }

    private static FormulaValue CopyStacked(List<ArrayValue> arrays, int rows, int cols, bool vertical)
    {
        var values = new CellValue[rows * cols];
        Array.Fill(values, CellValue.Error(ExcelError.NotAvailable));

        int offset = 0;
        foreach (ArrayValue array in arrays)
        {
            for (int r = 0; r < array.Rows; r++)
            {
                for (int c = 0; c < array.Columns; c++)
                {
                    int targetRow = vertical ? offset + r : r;
                    int targetCol = vertical ? c : offset + c;
                    values[(targetRow * cols) + targetCol] = NormalizeBlank(array[r, c]);
                }
            }

            offset += vertical ? array.Rows : array.Columns;
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static FormulaValue Flatten(EvalContext context, ReadOnlySpan<FormulaValue> args, bool toColumn)
    {
        if (args[0].IsError)
            return args[0];
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        int ignore = 0;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryInteger(context, args[1], out ignore, out CellValue error))
            return FormulaValue.Scalar(error);

        bool scanByColumn = false;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryBoolean(context, args[2], out scanByColumn, out error))
            return FormulaValue.Scalar(error);

        if (ignore is < 0 or > 3)
            return FormulaValue.Error(ExcelError.Value);

        List<CellValue> kept = CollectFlattened(array, ignore, scanByColumn);
        return kept.Count == 0
            ? FormulaValue.Error(ExcelError.Calc)
            : FormulaValue.Array(new ArrayValue(toColumn ? kept.Count : 1, toColumn ? 1 : kept.Count, [.. kept]));
    }

    private static List<CellValue> CollectFlattened(ArrayValue array, int ignore, bool scanByColumn)
    {
        var result = new List<CellValue>(array.Count);
        int outerLimit = scanByColumn ? array.Columns : array.Rows;
        int innerLimit = scanByColumn ? array.Rows : array.Columns;
        for (int outer = 0; outer < outerLimit; outer++)
        {
            for (int inner = 0; inner < innerLimit; inner++)
            {
                CellValue cell = scanByColumn ? array[inner, outer] : array[outer, inner];
                if (((ignore & 1) != 0 && cell.IsBlank) || ((ignore & 2) != 0 && cell.IsError))
                    continue;
                result.Add(NormalizeBlank(cell));
            }
        }

        return result;
    }

    private static FormulaValue Slice(EvalContext context, ReadOnlySpan<FormulaValue> args, bool isTake)
    {
        if (args[0].IsError)
            return args[0];
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        bool hasRows = !args[1].IsBlank;
        bool hasCols = args.Length > 2 && !args[2].IsBlank;
        if (!hasRows && !hasCols)
            return FormulaValue.Error(ExcelError.Value);

        int rowSpec = 0;
        if (hasRows && !Args.TryInteger(context, args[1], out rowSpec, out CellValue error))
            return FormulaValue.Scalar(error);

        int colSpec = 0;
        if (hasCols && !Args.TryInteger(context, args[2], out colSpec, out error))
            return FormulaValue.Scalar(error);

        if (isTake && ((hasRows && rowSpec == 0) || (hasCols && colSpec == 0)))
            return FormulaValue.Error(ExcelError.Value);

        (int startRow, int rowCount) = ResolveSliceBounds(array.Rows, hasRows ? rowSpec : null, isTake);
        (int startCol, int colCount) = ResolveSliceBounds(array.Columns, hasCols ? colSpec : null, isTake);
        if (rowCount <= 0 || colCount <= 0)
            return FormulaValue.Error(ExcelError.Calc);

        return CopySubArray(array, startRow, rowCount, startCol, colCount);
    }

    private static (int Start, int Count) ResolveSliceBounds(int total, int? spec, bool isTake)
    {
        if (spec is not { } count)
            return (0, total);

        if (isTake)
        {
            int taken = Math.Min(total, Math.Abs(count));
            return count > 0 ? (0, taken) : (total - taken, taken);
        }

        int dropped = Math.Min(total, Math.Abs(count));
        return count >= 0 ? (dropped, total - dropped) : (0, total - dropped);
    }

    private static FormulaValue CopySubArray(ArrayValue array, int startRow, int rowCount, int startCol, int colCount)
    {
        var values = new CellValue[rowCount * colCount];
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
                values[(r * colCount) + c] = NormalizeBlank(array[startRow + r, startCol + c]);
        }

        return FormulaValue.Array(new ArrayValue(rowCount, colCount, values));
    }

    private static FormulaValue Choose(EvalContext context, ReadOnlySpan<FormulaValue> args, bool byColumn)
    {
        if (args[0].IsError)
            return args[0];
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        int limit = byColumn ? array.Columns : array.Rows;
        var indices = new List<int>(args.Length - 1);
        foreach (ref readonly FormulaValue arg in args[1..])
        {
            foreach (CellValue raw in new ArgumentValues(arg, context))
            {
                if (!Args.TryInteger(context, FormulaValue.Scalar(raw), out int index, out CellValue error))
                    return FormulaValue.Scalar(error);
                if (index == 0 || Math.Abs(index) > limit)
                    return FormulaValue.Error(ExcelError.Value);
                indices.Add(index > 0 ? index - 1 : limit + index);
            }
        }

        return indices.Count == 0
            ? FormulaValue.Error(ExcelError.Value)
            : GatherChosen(array, indices, byColumn);
    }

    private static FormulaValue GatherChosen(ArrayValue array, List<int> indices, bool byColumn)
    {
        int rows = byColumn ? array.Rows : indices.Count;
        int cols = byColumn ? indices.Count : array.Columns;
        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        var values = new CellValue[rows * cols];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                CellValue cell = byColumn ? array[r, indices[c]] : array[indices[r], c];
                values[(r * cols) + c] = NormalizeBlank(cell);
            }
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static FormulaValue Wrap(EvalContext context, ReadOnlySpan<FormulaValue> args, bool byRows)
    {
        if (args[0].IsError)
            return args[0];
        if (!Operators.TryGetArray(args[0], context, out ArrayValue vector) || (vector.Rows > 1 && vector.Columns > 1))
            return FormulaValue.Error(ExcelError.Value);

        if (!Args.TryInteger(context, args[1], out int wrapCount, out CellValue error))
            return FormulaValue.Scalar(error);
        if (wrapCount < 1)
            return FormulaValue.Error(ExcelError.Number);

        CellValue padWith = args.Length > 2 && !args[2].IsBlank
            ? Operators.Dereference(args[2], context)
            : CellValue.Error(ExcelError.NotAvailable);

        int secondary = (vector.Count + wrapCount - 1) / wrapCount;
        int rows = byRows ? secondary : wrapCount;
        int cols = byRows ? wrapCount : secondary;
        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        return FillWrapped(vector, rows, cols, byRows, padWith);
    }

    private static FormulaValue FillWrapped(ArrayValue vector, int rows, int cols, bool byRows, CellValue padWith)
    {
        var values = new CellValue[rows * cols];
        Array.Fill(values, padWith);
        for (int i = 0; i < vector.Count; i++)
        {
            int r = byRows ? i / cols : i % rows;
            int c = byRows ? i % cols : i / rows;
            values[(r * cols) + c] = NormalizeBlank(vector.Values[i]);
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static FormulaValue Expand(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].IsError)
            return args[0];
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        int rows = array.Rows;
        if (!args[1].IsBlank && !Args.TryInteger(context, args[1], out rows, out CellValue error))
            return FormulaValue.Scalar(error);

        int cols = array.Columns;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryInteger(context, args[2], out cols, out error))
            return FormulaValue.Scalar(error);

        if (rows < array.Rows || cols < array.Columns)
            return FormulaValue.Error(ExcelError.Value);
        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        CellValue padWith = args.Length > 3 && !args[3].IsBlank
            ? Operators.Dereference(args[3], context)
            : CellValue.Error(ExcelError.NotAvailable);

        return FillExpanded(array, rows, cols, padWith);
    }

    private static FormulaValue FillExpanded(ArrayValue array, int rows, int cols, CellValue padWith)
    {
        var values = new CellValue[rows * cols];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                values[(r * cols) + c] = r < array.Rows && c < array.Columns
                    ? NormalizeBlank(array[r, c])
                    : padWith;
            }
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static CellValue NormalizeBlank(in CellValue value) =>
        value.IsBlank ? CellValue.Number(0) : value;
}
