namespace Cellwright.Formulas.Functions;

/// <summary>
/// Excel 2022–2024 substring extraction and splitting functions (<c>TEXTBEFORE</c>,
/// <c>TEXTAFTER</c>, <c>TEXTSPLIT</c>).
/// </summary>
internal static class TextSplitFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("TEXTBEFORE", 2, 6, (c, a) => ExtractAround(c, a, before: true), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("TEXTAFTER", 2, 6, (c, a) => ExtractAround(c, a, before: false), wantsRawReferences: true, acceptsErrors: true);
        builder.Add("TEXTSPLIT", 2, 6, TextSplit, wantsRawReferences: true, acceptsErrors: true);
    }

    private static FormulaValue ExtractAround(EvalContext context, ReadOnlySpan<FormulaValue> args, bool before)
    {
        CellValue source = Operators.Dereference(args[0], context);
        if (source.IsError)
            return FormulaValue.Scalar(source);

        if (!TryExtractDelimiters(args[1], context, out List<string> delimiters, out CellValue error))
            return FormulaValue.Scalar(error);

        if (!TryReadExtractOptions(context, args, out int instance, out StringComparison comparison, out bool matchEnd, out FormulaValue notFound, out error))
            return FormulaValue.Scalar(error);

        string text = Coercion.ToText(source, context.DateSystem);
        if (delimiters.All(static d => d.Length == 0))
            return FormulaValue.Text((before ? instance < 0 : instance > 0) ? text : string.Empty);

        return FindInstance(text, delimiters, instance, comparison, matchEnd, before, notFound);
    }

    private static bool TryReadExtractOptions(
        EvalContext context,
        ReadOnlySpan<FormulaValue> args,
        out int instance,
        out StringComparison comparison,
        out bool matchEnd,
        out FormulaValue notFound,
        out CellValue error)
    {
        instance = 1;
        comparison = StringComparison.Ordinal;
        matchEnd = false;
        notFound = args.Length > 5 && !args[5].IsBlank ? FormulaValue.Scalar(Operators.Dereference(args[5], context)) : FormulaValue.Error(ExcelError.NotAvailable);
        error = CellValue.Blank;

        if (args.Length > 2 && !args[2].IsBlank && !Args.TryInteger(context, args[2], out instance, out error))
            return false;

        int matchMode = 0;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryInteger(context, args[3], out matchMode, out error))
            return false;

        int endMode = 0;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryInteger(context, args[4], out endMode, out error))
            return false;

        if (instance == 0 || matchMode is not (0 or 1) || endMode is not (0 or 1))
        {
            error = CellValue.Error(ExcelError.Value);
            return false;
        }

        comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        matchEnd = endMode == 1;
        return true;
    }

    private static FormulaValue FindInstance(
        string text,
        List<string> delimiters,
        int instance,
        StringComparison comparison,
        bool matchEnd,
        bool before,
        FormulaValue notFound)
    {
        if (text.Length == 0 || Math.Abs(instance) > text.Length)
            return Math.Abs(instance) > Math.Max(1, text.Length) ? FormulaValue.Error(ExcelError.Value) : notFound;

        List<(int Index, int Length)> matches = CollectMatches(text, delimiters, comparison, fromRight: instance < 0);
        int target = Math.Abs(instance);
        if (target <= matches.Count)
        {
            (int index, int length) = matches[target - 1];
            return FormulaValue.Text(before ? text[..index] : text[(index + length)..]);
        }

        if (matchEnd && target == matches.Count + 1)
            return FormulaValue.Text((before ? instance > 0 : instance < 0) ? text : string.Empty);

        return notFound;
    }

    private static FormulaValue TextSplit(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue source = Operators.Dereference(args[0], context);
        if (source.IsError)
            return FormulaValue.Scalar(source);

        bool hasCol = !args[1].IsBlank;
        bool hasRow = args.Length > 2 && !args[2].IsBlank;
        if (!hasCol && !hasRow)
            return FormulaValue.Error(ExcelError.Value);

        List<string> colDelimiters = [];
        List<string> rowDelimiters = [];
        if ((hasCol && !TryExtractDelimiters(args[1], context, out colDelimiters, out CellValue error)) ||
            (hasRow && !TryExtractDelimiters(args[2], context, out rowDelimiters, out error)))
        {
            return FormulaValue.Scalar(error);
        }

        if (!TryReadSplitOptions(context, args, out bool ignoreEmpty, out StringComparison comparison, out CellValue padWith, out error))
            return FormulaValue.Scalar(error);

        string text = Coercion.ToText(source, context.DateSystem);
        if (text.Length == 0)
            return ignoreEmpty ? FormulaValue.Error(ExcelError.Calc) : FormulaValue.Array(ArrayValue.Single(CellValue.Text(string.Empty)));

        return BuildSplitGrid(text, colDelimiters, rowDelimiters, ignoreEmpty, comparison, padWith);
    }

    private static bool TryReadSplitOptions(
        EvalContext context,
        ReadOnlySpan<FormulaValue> args,
        out bool ignoreEmpty,
        out StringComparison comparison,
        out CellValue padWith,
        out CellValue error)
    {
        ignoreEmpty = false;
        comparison = StringComparison.Ordinal;
        padWith = CellValue.Error(ExcelError.NotAvailable);
        error = CellValue.Blank;

        if (args.Length > 3 && !args[3].IsBlank && !Args.TryBoolean(context, args[3], out ignoreEmpty, out error))
            return false;

        int matchMode = 0;
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryInteger(context, args[4], out matchMode, out error))
            return false;

        if (matchMode is not (0 or 1))
        {
            error = CellValue.Error(ExcelError.Value);
            return false;
        }

        if (args.Length > 5 && !args[5].IsBlank)
            padWith = Operators.Dereference(args[5], context);

        comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        return true;
    }

    private static FormulaValue BuildSplitGrid(
        string text,
        List<string> colDelimiters,
        List<string> rowDelimiters,
        bool ignoreEmpty,
        StringComparison comparison,
        CellValue padWith)
    {
        List<string> rawRows = rowDelimiters.Count > 0
            ? SplitByDelimiters(text, rowDelimiters, ignoreEmpty, comparison)
            : [text];

        var grid = new List<List<string>>(rawRows.Count);
        int maxCols = 0;
        foreach (string rowText in rawRows)
        {
            List<string> cols = colDelimiters.Count > 0
                ? SplitByDelimiters(rowText, colDelimiters, ignoreEmpty, comparison)
                : [rowText];
            if (ignoreEmpty && cols.Count == 0)
                continue;

            grid.Add(cols);
            if (cols.Count > maxCols)
                maxCols = cols.Count;
        }

        if (grid.Count == 0 || maxCols == 0)
            return FormulaValue.Error(ExcelError.Calc);

        return MaterializeSplitGrid(grid, maxCols, padWith);
    }

    private static FormulaValue MaterializeSplitGrid(List<List<string>> grid, int maxCols, CellValue padWith)
    {
        int rows = grid.Count;
        if ((long)rows * maxCols > ArrayValue.MaxCells)
            return FormulaValue.Error(ExcelError.Value);

        var values = new CellValue[rows * maxCols];
        for (int r = 0; r < rows; r++)
        {
            List<string> row = grid[r];
            for (int c = 0; c < maxCols; c++)
                values[(r * maxCols) + c] = c < row.Count ? CellValue.Text(row[c]) : padWith;
        }

        return FormulaValue.Array(new ArrayValue(rows, maxCols, values));
    }

    private static List<string> SplitByDelimiters(
        string text,
        List<string> delimiters,
        bool ignoreEmpty,
        StringComparison comparison)
    {
        List<(int Index, int Length)> matches = CollectMatches(text, delimiters, comparison, fromRight: false);
        var parts = new List<string>(matches.Count + 1);
        int start = 0;
        foreach ((int index, int length) in matches)
        {
            string segment = text[start..index];
            if (!ignoreEmpty || segment.Length > 0)
                parts.Add(segment);
            start = index + length;
        }

        string tail = text[start..];
        if (!ignoreEmpty || tail.Length > 0)
            parts.Add(tail);

        return parts;
    }

    private static List<(int Index, int Length)> CollectMatches(
        string text,
        List<string> delimiters,
        StringComparison comparison,
        bool fromRight)
    {
        var matches = new List<(int Index, int Length)>();
        int pos = fromRight ? text.Length : 0;
        while (fromRight ? pos >= 0 : pos <= text.Length)
        {
            if (!TryMatchAt(text, delimiters, pos, comparison, out int matchedLength))
            {
                pos += fromRight ? -1 : 1;
                continue;
            }

            matches.Add((pos, matchedLength));
            if (fromRight)
                pos--;
            else
                pos += Math.Max(1, matchedLength);
        }

        return matches;
    }

    private static bool TryMatchAt(
        string text,
        List<string> delimiters,
        int position,
        StringComparison comparison,
        out int length)
    {
        length = -1;
        foreach (string delimiter in delimiters)
        {
            if (delimiter.Length == 0 || position + delimiter.Length > text.Length)
                continue;
            if (text.AsSpan(position, delimiter.Length).Equals(delimiter.AsSpan(), comparison) &&
                delimiter.Length > length)
            {
                length = delimiter.Length;
            }
        }

        return length > 0;
    }

    private static bool TryExtractDelimiters(
        in FormulaValue arg,
        EvalContext context,
        out List<string> delimiters,
        out CellValue error)
    {
        delimiters = [];
        error = CellValue.Blank;
        foreach (CellValue item in new ArgumentValues(arg, context))
        {
            if (item.IsError)
            {
                error = item;
                return false;
            }

            delimiters.Add(Coercion.ToText(item, context.DateSystem));
        }

        return true;
    }
}
