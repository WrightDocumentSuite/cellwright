namespace Cellwright.Formulas.Functions;

/// <summary>
/// The dispatching aggregates (<c>SUBTOTAL</c>, <c>AGGREGATE</c>), the "A" variants that count
/// text and logicals, and the functions that inspect a cell's formula rather than its value.
/// Real workbooks lean on these heavily, which the corpus oracle made obvious.
/// </summary>
internal static class AggregateFunctions
{
    /// <summary>
    /// Function numbers shared by SUBTOTAL and AGGREGATE. SUBTOTAL stops at 11; the rest are
    /// AGGREGATE's own, and the last six take a second argument after the data.
    /// </summary>
    private static readonly string[] Dispatch =
    [
        "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PRODUCT",
        "STDEV", "STDEVP", "SUM", "VAR", "VARP",
        "MEDIAN", "MODE", "LARGE", "SMALL", "PERCENTILE.INC", "QUARTILE.INC",
        "PERCENTILE.EXC", "QUARTILE.EXC",
    ];

    /// <summary>Highest function number SUBTOTAL accepts; AGGREGATE goes further.</summary>
    private const int SubtotalMaxCode = 11;

    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add(
            "SUBTOTAL", 2, FunctionDefinition.Unlimited, Subtotal,
            wantsRawReferences: true, acceptsErrors: true, dependsOnVisibility: true);
        builder.Add(
            "AGGREGATE", 3, FunctionDefinition.Unlimited, Aggregate,
            wantsRawReferences: true, acceptsErrors: true, dependsOnVisibility: true);
        builder.Add("AVERAGEA", 1, FunctionDefinition.Unlimited, (c, a) => Anything(c, a, Mode.Average), wantsRawReferences: true);
        builder.Add("MAXA", 1, FunctionDefinition.Unlimited, (c, a) => Anything(c, a, Mode.Max), wantsRawReferences: true);
        builder.Add("MINA", 1, FunctionDefinition.Unlimited, (c, a) => Anything(c, a, Mode.Min), wantsRawReferences: true);
        builder.Add("FLOOR.MATH", 1, 3, (c, a) => MathStep(c, a, ceiling: false));
        builder.Add("CEILING.MATH", 1, 3, (c, a) => MathStep(c, a, ceiling: true));
        builder.Add("FLOOR.PRECISE", 1, 2, (c, a) => MathStep(c, a, ceiling: false));
        builder.Add("CEILING.PRECISE", 1, 2, (c, a) => MathStep(c, a, ceiling: true));
        builder.Add("ISO.CEILING", 1, 2, (c, a) => MathStep(c, a, ceiling: true));
        builder.Add("FORMULATEXT", 1, 1, FormulaText, wantsRawReferences: true, acceptsErrors: true);
        builder.Add("ISFORMULA", 1, 1, IsFormula, wantsRawReferences: true, acceptsErrors: true);
    }

    private enum Mode
    {
        Average,
        Max,
        Min,
    }

    /// <summary>
    /// <c>SUBTOTAL(code, …)</c> forwards to another aggregate. Codes above 100 skip rows
    /// hidden by hand; rows a filter hid are skipped by both code ranges, which is what makes
    /// SUBTOTAL(9, …) the total of what the user can see.
    /// </summary>
    private static FormulaValue Subtotal(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int code, out CellValue error))
            return FormulaValue.Scalar(error);

        bool skipHidden = code > 100;
        code = skipHidden ? code - 100 : code;
        if (code is < 1 or > SubtotalMaxCode)
            return FormulaValue.Error(ExcelError.Value);

        return Forward(context, code, args[1..], new Filter(false, skipHidden, true, SkipFiltered: true));
    }

    /// <summary>
    /// <c>AGGREGATE(code, options, …)</c> adds an options flag saying what to leave out:
    /// hidden rows, error values, nested aggregates, or nothing at all.
    /// </summary>
    private static FormulaValue Aggregate(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (!Args.TryInteger(context, args[0], out int code, out CellValue error) ||
            !Args.TryInteger(context, args[1], out int options, out error))
        {
            return FormulaValue.Scalar(error);
        }

        bool skipHidden = options is 1 or 3 or 5 or 7;
        var filter = new Filter(
            IgnoreErrors: options is 2 or 3 or 6 or 7,
            SkipHidden: skipHidden,
            IgnoreNested: options is >= 0 and <= 3,
            SkipFiltered: skipHidden);
        return Forward(context, code, args[2..], filter);
    }

    /// <summary>
    /// What a dispatching aggregate leaves out of its arguments. Hiding a row by hand and
    /// hiding it with a filter are separate reasons: SUBTOTAL always drops the second kind,
    /// while the first kind only goes with the 101-111 codes.
    /// </summary>
    private readonly record struct Filter(bool IgnoreErrors, bool SkipHidden, bool IgnoreNested, bool SkipFiltered);

    private static FormulaValue Forward(EvalContext context, int code, ReadOnlySpan<FormulaValue> args, Filter filter)
    {
        if (code < 1 || code > Dispatch.Length)
            return FormulaValue.Error(ExcelError.Value);

        FunctionDefinition? target = FunctionRegistry.Resolve(Dispatch[code - 1]);
        if (target is null)
            return FormulaValue.Error(ExcelError.Name);

        var filtered = new FormulaValue[args.Length];
        for (int i = 0; i < args.Length; i++)
        {
            filtered[i] = Prepare(context, args[i], filter);

            // The dispatcher accepts errors so that AGGREGATE can be told to ignore them; when
            // it was not, an error handed in whole propagates the way the evaluator would have
            // propagated it. The target is called directly and never gets the chance.
            if (!filter.IgnoreErrors && filtered[i].Kind == FormulaValueKind.Scalar && filtered[i].IsError)
                return filtered[i];
        }

        return target.Handler(context, filtered);
    }

    /// <summary>
    /// Materializes an argument for a dispatching aggregate. Cells that hold their own
    /// SUBTOTAL or AGGREGATE are dropped, which is how Excel keeps a grand total from counting
    /// the subtotals stacked above it; hidden rows and errors go the same way when asked for.
    /// </summary>
    private static FormulaValue Prepare(EvalContext context, in FormulaValue value, Filter filter)
    {
        if (value.Kind == FormulaValueKind.Scalar)
            return filter.IgnoreErrors && value.IsError ? FormulaValue.Blank : value;

        if (value.Kind != FormulaValueKind.Reference)
            return filter.IgnoreErrors ? WithoutErrors(context, value) : value;

        ReferenceValue reference = value.AsReference;
        var kept = new List<CellValue>();
        for (ReferenceValue? area = reference; area is not null; area = area.Next)
        {
            for (int row = area.FirstRow; row <= area.LastRow; row++)
            {
                if (IsSkippedRow(area.Sheet, row, filter))
                    continue;

                for (int column = area.FirstColumn; column <= area.LastColumn; column++)
                {
                    if (filter.IgnoreNested && IsNestedAggregate(area.Sheet, row, column))
                        continue;

                    CellValue cell = context.ReadCell(area.Sheet, row, column);
                    if (filter.IgnoreErrors && cell.IsError)
                        continue;
                    kept.Add(cell);
                }
            }
        }

        return FormulaValue.Array(new ArrayValue(1, kept.Count, [.. kept]));
    }

    /// <summary>
    /// Whether the row is left out. A hidden row inside the range of an active filter is
    /// treated as filtered away, because that is the only trace the file keeps of it.
    /// </summary>
    private static bool IsSkippedRow(Worksheet sheet, int row, Filter filter)
    {
        if (!filter.SkipHidden && !filter.SkipFiltered)
            return false;
        if (!sheet.Row(row).Hidden)
            return false;
        return filter.SkipHidden || sheet.IsFilteredAway(row);
    }

    private static bool IsNestedAggregate(Worksheet sheet, int row, int column)
    {
        string? formula = sheet.Formulas.Get(row, column);
        if (formula is null)
            return false;

        ReadOnlySpan<char> text = formula.AsSpan().TrimStart();
        return text.StartsWith("SUBTOTAL", StringComparison.OrdinalIgnoreCase) ||
               text.StartsWith("AGGREGATE", StringComparison.OrdinalIgnoreCase) ||
               text.StartsWith("_xlfn.AGGREGATE", StringComparison.OrdinalIgnoreCase);
    }

    private static FormulaValue WithoutErrors(EvalContext context, in FormulaValue value)
    {
        var kept = new List<CellValue>();
        foreach (CellValue item in new ArgumentValues(value, context))
        {
            if (!item.IsError)
                kept.Add(item);
        }

        return FormulaValue.Array(new ArrayValue(1, kept.Count, [.. kept]));
    }

    /// <summary>
    /// The "A" aggregates count text as zero and logicals as 0/1 instead of skipping them.
    /// </summary>
    private static FormulaValue Anything(EvalContext context, ReadOnlySpan<FormulaValue> args, Mode mode)
    {
        double sum = 0;
        double min = double.MaxValue;
        double max = double.MinValue;
        int count = 0;

        foreach (ref readonly FormulaValue arg in args)
        {
            // The A-suffixed aggregates count what the plain ones skip — text as zero, logicals
            // as one and nothing — but a logical written into an array constant is passed over
            // even by these, which is why AVERAGEA({1;2;4;"Text";TRUE}) divides by four.
            bool literalArray = arg.Kind == FormulaValueKind.Array;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (value.IsError)
                    return FormulaValue.Scalar(value);
                if (value.IsBlank || (literalArray && value.IsBoolean))
                    continue;

                double number = value.Kind switch
                {
                    CellValueKind.Number => value.GetNumber(),
                    CellValueKind.Boolean => value.GetBoolean() ? 1 : 0,
                    CellValueKind.DateTime or CellValueKind.TimeSpan => value.ToSerialNumber(context.DateSystem),
                    _ => 0,
                };

                sum += number;
                min = Math.Min(min, number);
                max = Math.Max(max, number);
                count++;
            }
        }

        return mode switch
        {
            Mode.Average => count == 0 ? FormulaValue.Error(ExcelError.DivisionByZero) : FormulaValue.Number(sum / count),
            Mode.Max => FormulaValue.Number(count == 0 ? 0 : max),
            _ => FormulaValue.Number(count == 0 ? 0 : min),
        };
    }

    /// <summary>
    /// <c>FLOOR.MATH</c> and <c>CEILING.MATH</c> default the step to one and, unlike the older
    /// pair, accept negative numbers by rounding toward or away from zero as asked.
    /// </summary>
    private static FormulaValue MathStep(EvalContext context, ReadOnlySpan<FormulaValue> args, bool ceiling)
    {
        if (!Args.TryNumber(context, args[0], out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        double significance = 1;
        if (args.Length > 1 && !args[1].IsBlank && !Args.TryNumber(context, args[1], out significance, out error))
            return FormulaValue.Scalar(error);

        double mode = 0;
        if (args.Length > 2 && !args[2].IsBlank && !Args.TryNumber(context, args[2], out mode, out error))
            return FormulaValue.Scalar(error);

        if (significance == 0 || number == 0)
            return FormulaValue.Number(0);

        significance = Math.Abs(significance);
        bool towardZero = number < 0 && mode != 0;
        double ratio = number / significance;
        double rounded = ceiling
            ? towardZero ? Math.Floor(ratio) : Math.Ceiling(ratio)
            : towardZero ? Math.Ceiling(ratio) : Math.Floor(ratio);

        return FormulaValue.Number(rounded * significance);
    }

    private static FormulaValue FormulaText(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.NotAvailable);

        ReferenceValue reference = args[0].AsReference;
        string? formula = reference.Sheet.Formulas.Get(reference.FirstRow, reference.FirstColumn);
        return formula is null ? FormulaValue.Error(ExcelError.NotAvailable) : FormulaValue.Text("=" + formula);
    }

    private static FormulaValue IsFormula(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].Kind != FormulaValueKind.Reference)
            return FormulaValue.Boolean(false);

        ReferenceValue reference = args[0].AsReference;
        return FormulaValue.Boolean(reference.Sheet.Formulas.Get(reference.FirstRow, reference.FirstColumn) is not null);
    }
}
