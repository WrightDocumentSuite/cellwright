namespace Cellwright.Formulas.Functions;

/// <summary>
/// Lookup and reference functions. These need the geometry of their arguments, so they take
/// references untouched rather than through implicit intersection.
/// </summary>
internal static class LookupFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("VLOOKUP", 3, 4, (c, a) => Lookup(c, a, vertical: true), wantsRawReferences: true);
        builder.Add("HLOOKUP", 3, 4, (c, a) => Lookup(c, a, vertical: false), wantsRawReferences: true);
        builder.Add("INDEX", 2, 3, Index, wantsRawReferences: true);
        builder.Add("MATCH", 2, 3, Match, wantsRawReferences: true);
        builder.Add("XLOOKUP", 3, 5, XLookup, wantsRawReferences: true);
        builder.Add("XMATCH", 2, 3, XMatch, wantsRawReferences: true);
        builder.Add("CHOOSE", 2, FunctionDefinition.Unlimited, Choose, wantsRawReferences: true);
        builder.Add("OFFSET", 3, 5, Offset, wantsRawReferences: true, isVolatile: true);
        builder.Add("INDIRECT", 1, 2, Indirect, isVolatile: true);
        builder.Add("ROW", 0, 1, (c, a) => Position(c, a, row: true), wantsRawReferences: true);
        builder.Add("COLUMN", 0, 1, (c, a) => Position(c, a, row: false), wantsRawReferences: true);
        builder.Add("ROWS", 1, 1, (c, a) => Size(c, a, rows: true), wantsRawReferences: true);
        builder.Add("COLUMNS", 1, 1, (c, a) => Size(c, a, rows: false), wantsRawReferences: true);
        builder.Add("ADDRESS", 2, 5, Address);
        builder.Add("HYPERLINK", 1, 2, Hyperlink);
    }

    /// <summary>
    /// <c>HYPERLINK(location, [friendly])</c>. To a calculation this is only what the cell
    /// shows: the jump belongs to whatever opens the file, and Excel keeps no record of the
    /// link in the sheet either — the formula is the link.
    /// </summary>
    private static FormulaValue Hyperlink(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        FormulaValue shown = args.Length > 1 && !args[1].IsBlank ? args[1] : args[0];
        return FormulaValue.Scalar(Operators.Dereference(shown, context));
    }

    private static FormulaValue Lookup(EvalContext context, ReadOnlySpan<FormulaValue> args, bool vertical)
    {
        CellValue key = Operators.Dereference(args[0], context);
        if (!Operators.TryGetArray(args[1], context, out ArrayValue table, out CachedRange? source))
            return FormulaValue.Error(ExcelError.Value);
        if (!Args.TryInteger(context, args[2], out int index, out CellValue error))
            return FormulaValue.Scalar(error);

        bool approximate = true;
        if (args.Length > 3 && !Args.TryBoolean(context, args[3], out approximate, out error))
            return FormulaValue.Scalar(error);

        // Excel splits the two failures: an index below one is a bad value, one past the
        // table's edge is a reference outside it.
        int limit = vertical ? table.Columns : table.Rows;
        if (index < 1)
            return FormulaValue.Error(ExcelError.Value);
        if (index > limit)
            return FormulaValue.Error(ExcelError.Reference);

        int lineCount = vertical ? table.Rows : table.Columns;
        LookupAxis axis = vertical ? LookupAxis.FirstColumn : LookupAxis.FirstRow;
        int found = approximate
            ? FindLine(context, table, key, lineCount, vertical, approximate: true)
            : FindExact(context, table, key, lineCount, vertical, source, axis);
        if (found < 0)
            return FormulaValue.Error(ExcelError.NotAvailable);

        return FormulaValue.Scalar(vertical ? table[found, index - 1] : table[index - 1, found]);
    }

    /// <summary>
    /// Exact match: answered from the range's hash index when there is one, otherwise by the
    /// same scan as before. The index is only consulted for keys and lines it can represent
    /// faithfully, so both paths always agree.
    /// </summary>
    private static int FindExact(
        EvalContext context, ArrayValue table, in CellValue key, int lineCount, bool vertical,
        CachedRange? source, LookupAxis axis)
    {
        if (CanUseIndex(source, key) &&
            source!.GetIndex(axis, context.DateSystem) is { } index)
        {
            return index.TryFind(key, context.DateSystem, out int position) && position < lineCount ? position : -1;
        }

        return FindLine(context, table, key, lineCount, vertical, approximate: false);
    }

    private static bool CanUseIndex(CachedRange? source, in CellValue key) =>
        source is { IsRetained: true } &&
        LookupIndex.CanSearch(key) &&
        !(key.IsText && Wildcard.ContainsWildcards(key.GetText()));

    /// <summary>Shared with <c>LOOKUP</c>, whose vector search is this exact scan.</summary>
    internal static int FindLine(
        EvalContext context, ArrayValue table, in CellValue key, int lineCount, bool vertical, bool approximate)
    {
        int best = -1;
        for (int i = 0; i < lineCount; i++)
        {
            CellValue candidate = vertical ? table[i, 0] : table[0, i];
            int order = CompareLookup(context, candidate, key);
            if (order == 0)
                return i;

            // Approximate matches take the largest value not exceeding the key, and the data
            // is assumed sorted, so the scan can stop once it passes the key.
            if (!approximate)
                continue;
            if (!SameOrderingClass(candidate, key))
                continue;
            if (order < 0)
                best = i;
            else
                break;
        }

        return best;
    }

    private static int CompareLookup(EvalContext context, in CellValue candidate, in CellValue key) =>
        key.IsText && candidate.IsText && Wildcard.ContainsWildcards(key.GetText())
            ? Wildcard.IsMatch(candidate.GetText(), key.GetText()) ? 0 : -1
            : Coercion.Compare(candidate, key, context.DateSystem);

    /// <summary>
    /// Whether two values belong to the same ordering class. An approximate lookup settles for
    /// the nearest value below the key, and Excel looks for it only among values of the key's
    /// own type: text sorts above every number, so without this a text key would happily settle
    /// on the last number in the column instead of answering <c>#N/A</c>.
    /// </summary>
    private static bool SameOrderingClass(in CellValue candidate, in CellValue key) =>
        OrderingClass(candidate) == OrderingClass(key);

    private static int OrderingClass(in CellValue value) => value.Kind switch
    {
        CellValueKind.Text => 1,
        CellValueKind.Boolean => 2,

        // An empty cell is nothing to settle on, so it stands apart from everything.
        CellValueKind.Blank => 3,
        _ => 0,
    };

    private static FormulaValue Index(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Operators.TryGetArray(args[0], context, out ArrayValue table))
            return FormulaValue.Error(ExcelError.Value);
        if (!Args.TryInteger(context, args[1], out int row, out CellValue error))
            return FormulaValue.Scalar(error);

        int column = 1;
        if (args.Length > 2 && !Args.TryInteger(context, args[2], out column, out error))
            return FormulaValue.Scalar(error);

        // With a single index the lone axis of a one-row source is the column, not the row.
        if (args.Length <= 2 && table.Rows == 1)
        {
            column = row;
            row = 1;
        }

        if (row < 0 || column < 0 || row > table.Rows || column > table.Columns)
            return FormulaValue.Error(ExcelError.Reference);
        if (row == 0 || column == 0)
            return FormulaValue.Error(ExcelError.Value);

        return FormulaValue.Scalar(table[row - 1, column - 1]);
    }

    private static FormulaValue Match(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue key = Operators.Dereference(args[0], context);
        if (!Operators.TryGetArray(args[1], context, out ArrayValue table, out CachedRange? source))
            return FormulaValue.Error(ExcelError.Value);

        int type = 1;
        if (args.Length > 2 && !Args.TryInteger(context, args[2], out type, out CellValue error))
            return FormulaValue.Scalar(error);

        if (type == 0 && FindInVector(context, key, source) is { } indexed)
            return indexed >= 0 ? FormulaValue.Number(indexed + 1) : FormulaValue.Error(ExcelError.NotAvailable);

        int best = -1;
        for (int i = 0; i < table.Count; i++)
        {
            CellValue candidate = table.Values[i];
            int order = CompareLookup(context, candidate, key);
            if (order == 0)
                return FormulaValue.Number(i + 1);
            if (!SameOrderingClass(candidate, key))
                continue;
            if (type == 1 && order < 0)
                best = i;
            else if (type == -1 && order > 0)
                best = i;
        }

        return best < 0 ? FormulaValue.Error(ExcelError.NotAvailable) : FormulaValue.Number(best + 1);
    }

    private static FormulaValue XLookup(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue key = Operators.Dereference(args[0], context);
        if (!Operators.TryGetArray(args[1], context, out ArrayValue lookup, out CachedRange? source) ||
            !Operators.TryGetArray(args[2], context, out ArrayValue result))
        {
            return FormulaValue.Error(ExcelError.Value);
        }

        if (FindInVector(context, key, source) is { } indexed)
        {
            if (indexed < 0)
                return args.Length > 3 ? args[3] : FormulaValue.Error(ExcelError.NotAvailable);
            return FormulaValue.Scalar(indexed < result.Count ? result.Values[indexed] : CellValue.Error(ExcelError.Reference));
        }

        for (int i = 0; i < lookup.Count; i++)
        {
            if (CompareLookup(context, lookup.Values[i], key) != 0)
                continue;

            // The result range mirrors the lookup range's orientation.
            return FormulaValue.Scalar(i < result.Count ? result.Values[i] : CellValue.Error(ExcelError.Reference));
        }

        return args.Length > 3 ? args[3] : FormulaValue.Error(ExcelError.NotAvailable);
    }

    private static FormulaValue XMatch(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue key = Operators.Dereference(args[0], context);
        if (!Operators.TryGetArray(args[1], context, out ArrayValue lookup, out CachedRange? source))
            return FormulaValue.Error(ExcelError.Value);

        if (FindInVector(context, key, source) is { } indexed)
            return indexed >= 0 ? FormulaValue.Number(indexed + 1) : FormulaValue.Error(ExcelError.NotAvailable);

        for (int i = 0; i < lookup.Count; i++)
        {
            if (CompareLookup(context, lookup.Values[i], key) == 0)
                return FormulaValue.Number(i + 1);
        }

        return FormulaValue.Error(ExcelError.NotAvailable);
    }

    /// <summary>
    /// Position of an exact match through the range's hash index, or <see langword="null"/>
    /// when there is no usable index and the caller must scan.
    /// </summary>
    private static int? FindInVector(EvalContext context, in CellValue key, CachedRange? source)
    {
        if (!CanUseIndex(source, key))
            return null;
        if (source!.GetIndex(LookupAxis.AllValues, context.DateSystem) is not { } index)
            return null;

        return index.TryFind(key, context.DateSystem, out int position) ? position : -1;
    }

    private static FormulaValue Choose(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int index, out CellValue error))
            return FormulaValue.Scalar(error);
        if (index < 1 || index >= args.Length)
            return FormulaValue.Error(ExcelError.Value);
        return args[index];
    }

    private static FormulaValue Offset(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.Reference);

        ReferenceValue origin = args[0].AsReference;
        if (!Args.TryInteger(context, args[1], out int rowOffset, out CellValue error) ||
            !Args.TryInteger(context, args[2], out int columnOffset, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int height = origin.RowCount;
        int width = origin.ColumnCount;
        if (args.Length > 3 && !args[3].IsBlank && !Args.TryInteger(context, args[3], out height, out error))
            return FormulaValue.Scalar(error);
        if (args.Length > 4 && !args[4].IsBlank && !Args.TryInteger(context, args[4], out width, out error))
            return FormulaValue.Scalar(error);

        int firstRow = origin.FirstRow + rowOffset;
        int firstColumn = origin.FirstColumn + columnOffset;
        if (height <= 0 || width <= 0 || firstRow < 1 || firstColumn < 1 ||
            !SheetLimits.IsValidRow(firstRow + height - 1) || !SheetLimits.IsValidColumn(firstColumn + width - 1))
        {
            return FormulaValue.Error(ExcelError.Reference);
        }

        return FormulaValue.Reference(new ReferenceValue(
            origin.Sheet, firstRow, firstColumn, firstRow + height - 1, firstColumn + width - 1));
    }

    private static FormulaValue Indirect(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string text = Args.Text(context, args[0]);
        if (text.Length == 0)
            return FormulaValue.Error(ExcelError.Reference);

        // A reference longer than this is refused before it is parsed (MS-OI29500 2.1.974).
        // No real reference approaches the limit, and it caps what a hostile formula can ask
        // the parser to chew on.
        const int LongestReference = 255;
        if (text.Length > LongestReference)
            return FormulaValue.Error(ExcelError.Value);

        // Parsed once per distinct target text: an INDIRECT filled down a column repeats the
        // same handful of strings on every row.
        Parsing.FormulaProgram? program = context.GetProgram(text);
        if (program is null)
            return FormulaValue.Error(ExcelError.Reference);

        FormulaValue value = context.EvaluateNested(program);
        return value.IsReference ? value : FormulaValue.Error(ExcelError.Reference);
    }

    private static FormulaValue Position(EvalContext context, ReadOnlySpan<FormulaValue> args, bool row)
    {
        if (args.Length == 0 || args[0].IsBlank)
            return FormulaValue.Number(row ? context.Cell.Row : context.Cell.Column);
        if (args[0].Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.Reference);

        ReferenceValue reference = args[0].AsReference;
        return FormulaValue.Number(row ? reference.FirstRow : reference.FirstColumn);
    }

    private static FormulaValue Size(EvalContext context, ReadOnlySpan<FormulaValue> args, bool rows)
    {
        if (args[0].Kind == FormulaValueKind.Reference)
        {
            ReferenceValue reference = args[0].AsReference.Bounds();
            return FormulaValue.Number(rows ? reference.RowCount : reference.ColumnCount);
        }

        if (args[0].IsArray)
            return FormulaValue.Number(rows ? args[0].AsArray.Rows : args[0].AsArray.Columns);

        return FormulaValue.Number(1);
    }

    private static FormulaValue Address(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int row, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int column, out error))
        {
            return FormulaValue.Scalar(error);
        }

        int kind = 1;
        if (args.Length > 2 && !Args.TryInteger(context, args[2], out kind, out error))
            return FormulaValue.Scalar(error);
        if (!SheetLimits.IsValidRow(row) || !SheetLimits.IsValidColumn(column) || kind is < 1 or > 4)
            return FormulaValue.Error(ExcelError.Value);

        string rowAnchor = kind is 1 or 3 ? "$" : "";
        string columnAnchor = kind is 1 or 2 ? "$" : "";
        string address = $"{columnAnchor}{ColumnName.Format(column)}{rowAnchor}{row}";
        if (args.Length > 4)
            address = $"{Args.Text(context, args[4])}!{address}";

        return FormulaValue.Text(address);
    }
}
