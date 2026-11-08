namespace Cellwright.Formulas.Functions;

/// <summary>
/// Higher-order <c>LAMBDA</c> helper functions:
/// <c>ISOMITTED</c>, <c>MAP</c>, <c>REDUCE</c>, <c>SCAN</c>, <c>MAKEARRAY</c>, <c>BYROW</c>, <c>BYCOL</c>.
/// </summary>
internal static class LambdaFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("ISOMITTED", 1, 1, IsOmitted, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("MAP", 2, FunctionDefinition.Unlimited, Map, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("REDUCE", 2, 3, Reduce, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("SCAN", 2, 3, Scan, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("MAKEARRAY", 3, 3, MakeArray, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("BYROW", 2, 2, (c, a) => ByDimension(c, a, byColumn: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("BYCOL", 2, 2, (c, a) => ByDimension(c, a, byColumn: true), wantsRawReferences: true, acceptsErrors: true);
    }

    private static FormulaValue IsOmitted(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        FormulaValue.Boolean(args[0].IsOmitted);

    private static FormulaValue Map(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!args[^1].IsLambda)
            return args[^1].IsError ? args[^1] : FormulaValue.Error(ExcelError.Value);

        LambdaValue lambda = args[^1].AsLambda;
        int inputCount = args.Length - 1;
        if (lambda.Definition.ParameterSlots.Length < inputCount)
            return FormulaValue.Error(ExcelError.Value);

        var arrays = new ArrayValue[inputCount];
        int rows = 1;
        int cols = 1;
        for (int i = 0; i < inputCount; i++)
        {
            if (!Operators.TryGetArray(args[i], context, out ArrayValue arr))
                return FormulaValue.Error(ExcelError.Value);
            arrays[i] = arr;
            rows = Math.Max(rows, arr.Rows);
            cols = Math.Max(cols, arr.Columns);
        }

        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        return ExecuteMap(context, lambda, arrays, rows, cols);
    }

    private static FormulaValue ExecuteMap(
        EvalContext context, LambdaValue lambda, ArrayValue[] arrays, int rows, int cols)
    {
        var values = new CellValue[rows * cols];
        var callArgs = new FormulaValue[arrays.Length];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                for (int i = 0; i < arrays.Length; i++)
                    callArgs[i] = FormulaValue.Scalar(ReadBroadcast(arrays[i], r, c));

                FormulaValue res = context.InvokeLambda(lambda, callArgs);
                if (res.IsArray && res.AsArray.Count > 1)
                    values[(r * cols) + c] = CellValue.Error(ExcelError.Calc);
                else
                    values[(r * cols) + c] = NormalizeBlank(Operators.Dereference(res, context));
            }
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static FormulaValue Reduce(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryUnpackFoldArgs(context, args, out bool hasInitial, out FormulaValue initial, out ArrayValue array, out LambdaValue lambda, out FormulaValue error))
            return error;

        if (!hasInitial && array.Count == 0)
            return FormulaValue.Error(ExcelError.Calc);

        FormulaValue acc = hasInitial ? initial : FormulaValue.Scalar(NormalizeBlank(array.Values[0]));
        int start = hasInitial ? 0 : 1;
        for (int i = start; i < array.Count; i++)
            acc = context.InvokeLambda(lambda, [acc, FormulaValue.Scalar(NormalizeBlank(array.Values[i]))]);

        return acc;
    }

    private static FormulaValue Scan(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!TryUnpackFoldArgs(context, args, out bool hasInitial, out FormulaValue initial, out ArrayValue array, out LambdaValue lambda, out FormulaValue error))
            return error;

        var results = new CellValue[array.Count];
        FormulaValue acc = initial;
        for (int i = 0; i < array.Count; i++)
        {
            acc = i == 0 && !hasInitial
                ? FormulaValue.Scalar(NormalizeBlank(array.Values[0]))
                : context.InvokeLambda(lambda, [acc, FormulaValue.Scalar(NormalizeBlank(array.Values[i]))]);

            if (acc.IsArray && acc.AsArray.Count > 1)
                return FormulaValue.Error(ExcelError.Calc);

            results[i] = NormalizeBlank(Operators.Dereference(acc, context));
        }

        return FormulaValue.Array(new ArrayValue(array.Rows, array.Columns, results));
    }

    private static bool TryUnpackFoldArgs(
        EvalContext context,
        ReadOnlySpan<FormulaValue> args,
        out bool hasInitial,
        out FormulaValue initial,
        out ArrayValue array,
        out LambdaValue lambda,
        out FormulaValue error)
    {
        hasInitial = args.Length == 3 && !args[0].IsOmitted;
        initial = hasInitial ? FormulaValue.Scalar(NormalizeBlank(Operators.Dereference(args[0], context))) : FormulaValue.Blank;
        FormulaValue arrayArg = args.Length == 3 ? args[1] : args[0];
        FormulaValue lambdaArg = args[^1];
        array = null!;
        lambda = null!;

        if (!lambdaArg.IsLambda)
        {
            error = lambdaArg.IsError ? lambdaArg : FormulaValue.Error(ExcelError.Value);
            return false;
        }

        lambda = lambdaArg.AsLambda;
        if (lambda.Definition.ParameterSlots.Length < 2 || !Operators.TryGetArray(arrayArg, context, out array))
        {
            error = FormulaValue.Error(ExcelError.Value);
            return false;
        }

        error = default;
        return true;
    }

    private static FormulaValue MakeArray(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int rows, out CellValue err) ||
            !Args.TryInteger(context, args[1], out int cols, out err))
        {
            return FormulaValue.Scalar(err);
        }

        if (rows < 1 || cols < 1 || !args[2].IsLambda)
            return args[2].IsError ? args[2] : FormulaValue.Error(ExcelError.Value);

        LambdaValue lambda = args[2].AsLambda;
        if (lambda.Definition.ParameterSlots.Length < 2)
            return FormulaValue.Error(ExcelError.Value);
        if ((long)rows * cols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Number);

        var values = new CellValue[rows * cols];
        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= cols; c++)
            {
                FormulaValue res = context.InvokeLambda(lambda, [FormulaValue.Number(r), FormulaValue.Number(c)]);
                if (res.IsArray && res.AsArray.Count > 1)
                    return FormulaValue.Error(ExcelError.Calc);
                values[((r - 1) * cols) + (c - 1)] = NormalizeBlank(Operators.Dereference(res, context));
            }
        }

        return FormulaValue.Array(new ArrayValue(rows, cols, values));
    }

    private static FormulaValue ByDimension(EvalContext context, ReadOnlySpan<FormulaValue> args, bool byColumn)
    {
        if (args[0].IsError)
            return args[0];
        if (!args[1].IsLambda)
            return args[1].IsError ? args[1] : FormulaValue.Error(ExcelError.Value);

        LambdaValue lambda = args[1].AsLambda;
        if (lambda.Definition.ParameterSlots.Length < 1 || !Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        int count = byColumn ? array.Columns : array.Rows;
        int sliceLen = byColumn ? array.Rows : array.Columns;
        var results = new CellValue[count];

        for (int i = 0; i < count; i++)
        {
            var sliceValues = new CellValue[sliceLen];
            for (int j = 0; j < sliceLen; j++)
                sliceValues[j] = NormalizeBlank(byColumn ? array[j, i] : array[i, j]);

            var slice = new ArrayValue(byColumn ? sliceLen : 1, byColumn ? 1 : sliceLen, sliceValues);
            FormulaValue res = context.InvokeLambda(lambda, [FormulaValue.Array(slice)]);
            if (res.IsArray && res.AsArray.Count > 1)
                return FormulaValue.Error(ExcelError.Calc);

            results[i] = NormalizeBlank(Operators.Dereference(res, context));
        }

        return FormulaValue.Array(new ArrayValue(byColumn ? 1 : count, byColumn ? count : 1, results));
    }

    private static CellValue ReadBroadcast(ArrayValue array, int row, int col)
    {
        if (row < array.Rows && col < array.Columns)
            return NormalizeBlank(array[row, col]);
        if (array.Rows == 1 && array.Columns == 1)
            return NormalizeBlank(array[0, 0]);
        return CellValue.Error(ExcelError.NotAvailable);
    }

    private static CellValue NormalizeBlank(in CellValue value) =>
        value.IsBlank ? CellValue.Number(0) : value;
}
