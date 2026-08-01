namespace Cellwright.Formulas.Functions;

/// <summary>
/// <c>LOOKUP</c> — the elder of VLOOKUP and HLOOKUP, still everywhere in real files — and
/// <c>AREAS</c>, which counts the rectangles of a reference.
/// </summary>
internal static class VectorLookupFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("LOOKUP", 2, 3, Lookup, wantsRawReferences: true);
        builder.Add("AREAS", 1, 1, Areas, wantsRawReferences: true);
    }

    /// <summary>
    /// The vector form searches one row or column and answers from another; the array form
    /// searches the first column of a tall array (or the first row of a wide one) and answers
    /// from the last. Both searches are approximate: the largest value not past the key, among
    /// values of the key's own type.
    /// </summary>
    private static FormulaValue Lookup(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue key = Operators.Dereference(args[0], context);
        if (key.IsError)
            return FormulaValue.Scalar(key);
        if (!Operators.TryGetArray(args[1], context, out ArrayValue source))
            return FormulaValue.Error(ExcelError.Value);

        if (args.Length > 2 && !args[2].IsBlank)
            return VectorLookup(context, key, source, args[2]);

        // Array form: a wide array searches its first row and answers from its last; anything
        // square or tall searches the first column and answers from the last.
        bool byRow = source.Columns > source.Rows;
        int found = LookupFunctions.FindLine(
            context, source, key, byRow ? source.Columns : source.Rows, vertical: !byRow, approximate: true);
        if (found < 0)
            return FormulaValue.Error(ExcelError.NotAvailable);

        return FormulaValue.Scalar(byRow ? source[source.Rows - 1, found] : source[found, source.Columns - 1]);
    }

    private static FormulaValue VectorLookup(
        EvalContext context, in CellValue key, ArrayValue lookup, in FormulaValue resultArg)
    {
        if (lookup.Rows != 1 && lookup.Columns != 1)
            return FormulaValue.Error(ExcelError.NotAvailable);
        if (!Operators.TryGetArray(resultArg, context, out ArrayValue result))
            return FormulaValue.Error(ExcelError.Value);
        if (result.Rows != 1 && result.Columns != 1)
            return FormulaValue.Error(ExcelError.NotAvailable);

        bool vertical = lookup.Columns == 1;
        int found = LookupFunctions.FindLine(
            context, lookup, key, vertical ? lookup.Rows : lookup.Columns, vertical, approximate: true);
        if (found < 0)
            return FormulaValue.Error(ExcelError.NotAvailable);

        // The result vector follows its own orientation, however the lookup vector ran.
        return found < result.Count
            ? FormulaValue.Scalar(result.Values[found])
            : FormulaValue.Error(ExcelError.Reference);
    }

    /// <summary>How many rectangles the reference names (§18.17.7.10).</summary>
    private static FormulaValue Areas(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.Value);

        int count = 0;
        for (ReferenceValue? area = args[0].AsReference; area is not null; area = area.Next)
            count++;
        return FormulaValue.Number(count);
    }
}
