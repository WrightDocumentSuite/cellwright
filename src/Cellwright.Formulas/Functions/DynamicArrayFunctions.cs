using System.Text;

namespace Cellwright.Formulas.Functions;

/// <summary>
/// The dynamic array functions that reorder or select from an array they were handed.
/// </summary>
/// <remarks>
/// In Excel these spill across the grid; here the result stays an array value, so a formula
/// built around one — <c>SUM(FILTER(…))</c>, <c>INDEX(SORT(…),2,1)</c> — works, while a cell
/// holding one on its own shows the first element. That is the same bargain array formulas
/// struck before spilling existed, and the engine has kept to it since v1.
/// </remarks>
internal static class DynamicArrayFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("SORT", 1, 4, Sort, wantsRawReferences: true);
        builder.Add("SORTBY", 2, FunctionDefinition.Unlimited, SortBy, wantsRawReferences: true);
        builder.Add("FILTER", 2, 3, Filter, wantsRawReferences: true);
        builder.Add("UNIQUE", 1, 3, Unique, wantsRawReferences: true);
        builder.Add("TRANSPOSE", 1, 1, Transpose, wantsRawReferences: true);

        // The @ of a modern formula. Narrowing a range to the one cell that lines up with the
        // formula is what every scalar argument already gets, so this only has to ask for it.
        builder.Add(
            "SINGLE", 1, 1,
            static (c, a) => FormulaValue.Scalar(Operators.Dereference(a[0], c)),
            wantsRawReferences: true);
    }

    /// <summary>
    /// <c>SORT(array, [index], [order], [by_column])</c>. Rows are sorted by one of their
    /// columns, or columns by one of their rows, and equal entries keep the order they were in.
    /// </summary>
    private static FormulaValue Sort(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        int index = 1;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryInteger(context, args[1], out index, out CellValue error))
            return FormulaValue.Scalar(error);

        int order = 1;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryInteger(context, args[2], out order, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        bool byColumn = false;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryBoolean(context, args[3], out byColumn, out CellValue refused))
            return FormulaValue.Scalar(refused);

        int depth = byColumn ? array.Rows : array.Columns;
        if (order is not (1 or -1) || index < 1 || index > depth)
            return FormulaValue.Error(ExcelError.Value);

        int[] lanes = Identity(byColumn ? array.Columns : array.Rows);
        int key = index - 1;
        Stable(lanes, (a, b) => order * Coercion.Compare(
            Lane(array, byColumn, a, key), Lane(array, byColumn, b, key), context.DateSystem));

        return FormulaValue.Array(Gather(array, lanes, byColumn));
    }

    /// <summary>
    /// <c>SORTBY(array, by1, [order1], …)</c>. The keys are separate arrays, so the column
    /// sorted on need not be in the result — which is the whole reason this exists beside SORT.
    /// </summary>
    private static FormulaValue SortBy(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        var keys = new List<(ArrayValue By, int Order)>();
        bool byColumn = false;
        for (int i = 1; i < args.Length; i += 2)
        {
            if (!Operators.TryGetArray(args[i], context, out ArrayValue by))
                return FormulaValue.Error(ExcelError.Value);

            int order = 1;
            if (i + 1 < args.Length && !args[i + 1].IsBlank &&
                !Args.TryInteger(context, args[i + 1], out order, out CellValue error))
            {
                return FormulaValue.Scalar(error);
            }

            // A key written as a row sorts columns and one written as a column sorts rows; the
            // first key settles it and the rest have to agree.
            if (keys.Count == 0)
                byColumn = by.Rows == 1 && by.Columns > 1;

            if (order is not (1 or -1) || by.Count != (byColumn ? array.Columns : array.Rows))
                return FormulaValue.Error(ExcelError.Value);

            keys.Add((by, order));
        }

        int[] lanes = Identity(byColumn ? array.Columns : array.Rows);
        Stable(lanes, (a, b) => Rank(keys, a, b, context));
        return FormulaValue.Array(Gather(array, lanes, byColumn));
    }

    private static int Rank(List<(ArrayValue By, int Order)> keys, int left, int right, EvalContext context)
    {
        foreach ((ArrayValue by, int order) in keys)
        {
            int result = order * Coercion.Compare(by.Values[left], by.Values[right], context.DateSystem);
            if (result != 0)
                return result;
        }

        return 0;
    }

    /// <summary>
    /// <c>FILTER(array, include, [if_empty])</c>. A column of flags picks rows and a row of
    /// them picks columns; which it is follows from the shape rather than from an argument.
    /// </summary>
    private static FormulaValue Filter(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array) ||
            !Operators.TryGetArray(args[1], context, out ArrayValue include))
        {
            return FormulaValue.Error(ExcelError.Value);
        }

        bool byColumn = include.Rows == 1 && include.Columns > 1;
        int count = byColumn ? array.Columns : array.Rows;
        if (include.Count != count)
            return FormulaValue.Error(ExcelError.Value);

        var kept = new List<int>();
        for (int lane = 0; lane < count; lane++)
        {
            CellValue flag = include.Values[lane];
            if (flag.IsError)
                return FormulaValue.Scalar(flag);
            if (Truthy(flag))
                kept.Add(lane);
        }

        // Nothing left is not an error in itself, but it is nothing to show, so the caller gets
        // whatever it said to show instead — and #CALC! when it said nothing.
        if (kept.Count == 0)
            return args.Length > 2 ? args[2] : FormulaValue.Error(ExcelError.Calc);

        return FormulaValue.Array(Gather(array, [.. kept], byColumn));
    }

    /// <summary>
    /// <c>UNIQUE(array, [by_column], [exactly_once])</c>. Rows count as the same when every
    /// value in them matches, and text matches without regard to case, as it does everywhere else.
    /// </summary>
    private static FormulaValue Unique(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        bool byColumn = false;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryBoolean(context, args[1], out byColumn, out CellValue error))
            return FormulaValue.Scalar(error);

        bool once = false;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryBoolean(context, args[2], out once, out CellValue rejected))
            return FormulaValue.Scalar(rejected);

        int count = byColumn ? array.Columns : array.Rows;
        int depth = byColumn ? array.Rows : array.Columns;
        var keys = new string[count];
        var tally = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int lane = 0; lane < count; lane++)
        {
            keys[lane] = Signature(array, byColumn, lane, depth, context);
            tally[keys[lane]] = tally.GetValueOrDefault(keys[lane]) + 1;
        }

        var kept = new List<int>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int lane = 0; lane < count; lane++)
        {
            bool wanted = once ? tally[keys[lane]] == 1 : seen.Add(keys[lane]);
            if (wanted)
                kept.Add(lane);
        }

        return kept.Count == 0
            ? FormulaValue.Error(ExcelError.Calc)
            : FormulaValue.Array(Gather(array, [.. kept], byColumn));
    }

    private static FormulaValue Transpose(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue array))
            return FormulaValue.Error(ExcelError.Value);

        var values = new CellValue[array.Count];
        for (int row = 0; row < array.Rows; row++)
        {
            for (int column = 0; column < array.Columns; column++)
                values[(column * array.Rows) + row] = array[row, column];
        }

        return FormulaValue.Array(new ArrayValue(array.Columns, array.Rows, values));
    }

    /// <summary>What one row, or one column, holds — in a form two of them can be compared by.</summary>
    private static string Signature(ArrayValue array, bool byColumn, int lane, int depth, EvalContext context)
    {
        var builder = new StringBuilder();
        for (int i = 0; i < depth; i++)
        {
            CellValue value = Lane(array, byColumn, lane, i);

            // The kind travels with the text so that a number and its spelling stay apart.
            builder.Append((char)('a' + (int)value.Kind));
            builder.Append(value.IsText
                ? value.GetText().ToUpperInvariant()
                : Coercion.ToText(value, context.DateSystem));
            builder.Append('\u001f');
        }

        return builder.ToString();
    }

    private static CellValue Lane(ArrayValue array, bool byColumn, int lane, int depth) =>
        byColumn ? array[depth, lane] : array[lane, depth];

    private static bool Truthy(in CellValue value) => value.Kind switch
    {
        CellValueKind.Boolean => value.GetBoolean(),
        CellValueKind.Number => value.GetNumber() != 0,
        CellValueKind.DateTime or CellValueKind.TimeSpan => true,
        _ => false,
    };

    private static int[] Identity(int count)
    {
        var positions = new int[count];
        for (int i = 0; i < count; i++)
            positions[i] = i;

        return positions;
    }

    /// <summary>
    /// Sorts positions, breaking every tie by the position itself. That turns an unstable sort
    /// into a stable one, which matters here: equal rows must keep the order they were written in.
    /// </summary>
    private static void Stable(int[] positions, Comparison<int> compare) =>
        Array.Sort(positions, (a, b) => compare(a, b) is int order and not 0 ? order : a.CompareTo(b));

    /// <summary>Builds an array from the chosen rows, or the chosen columns, in the order given.</summary>
    private static ArrayValue Gather(ArrayValue array, int[] lanes, bool byColumn)
    {
        int rows = byColumn ? array.Rows : lanes.Length;
        int columns = byColumn ? lanes.Length : array.Columns;
        var values = new CellValue[rows * columns];
        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                values[(row * columns) + column] = byColumn
                    ? array[row, lanes[column]]
                    : array[lanes[row], column];
            }
        }

        return new ArrayValue(rows, columns, values);
    }
}
