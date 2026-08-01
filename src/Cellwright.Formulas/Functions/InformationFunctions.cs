namespace Cellwright.Formulas.Functions;

/// <summary>
/// Type predicates and conversions. These are the functions that must see errors instead of
/// propagating them, which is what <see cref="FunctionDefinition.AcceptsErrors"/> enables.
/// </summary>
internal static class InformationFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("ISBLANK", 1, 1, (c, a) => Test(c, a[0], static v => v.IsBlank), acceptsErrors: true);
        builder.Add("ISNUMBER", 1, 1, (c, a) => Test(c, a[0], static v => v.IsNumber || v.IsDateTime || v.IsTimeSpan), acceptsErrors: true);
        builder.Add("ISTEXT", 1, 1, (c, a) => Test(c, a[0], static v => v.IsText), acceptsErrors: true);
        builder.Add("ISNONTEXT", 1, 1, (c, a) => Test(c, a[0], static v => !v.IsText), acceptsErrors: true);
        builder.Add("ISLOGICAL", 1, 1, (c, a) => Test(c, a[0], static v => v.IsBoolean), acceptsErrors: true);
        builder.Add("ISERROR", 1, 1, (c, a) => Test(c, a[0], static v => v.IsError), acceptsErrors: true);
        builder.Add("ISERR", 1, 1, (c, a) => Test(c, a[0], static v => v.IsError && v.GetError() != ExcelError.NotAvailable), acceptsErrors: true);
        builder.Add("ISNA", 1, 1, (c, a) => Test(c, a[0], static v => v.IsError && v.GetError() == ExcelError.NotAvailable), acceptsErrors: true);
        builder.Add("ISEVEN", 1, 1, (c, a) => Parity(c, a[0], even: true));
        builder.Add("ISODD", 1, 1, (c, a) => Parity(c, a[0], even: false));
        builder.Add("ISREF", 1, 1, static (_, a) => FormulaValue.Boolean(a[0].IsReference), acceptsErrors: true, wantsRawReferences: true);
        builder.Add("N", 1, 1, N, acceptsErrors: true);
        builder.Add("NA", 0, 0, static (_, _) => FormulaValue.Error(ExcelError.NotAvailable));
        builder.Add("TYPE", 1, 1, Type, acceptsErrors: true);
        builder.Add("ERROR.TYPE", 1, 1, ErrorType, acceptsErrors: true);
    }

    private static FormulaValue Test(EvalContext context, in FormulaValue value, Func<CellValue, bool> predicate) =>
        FormulaValue.Boolean(predicate(Operators.Dereference(value, context)));

    private static FormulaValue Parity(EvalContext context, in FormulaValue value, bool even)
    {
        if (!Args.TryNumber(context, value, out double number, out CellValue error))
            return FormulaValue.Scalar(error);

        bool isEven = Math.Abs(Math.Truncate(number)) % 2 == 0;
        return FormulaValue.Boolean(isEven == even);
    }

    private static FormulaValue N(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue value = Operators.Dereference(args[0], context);
        return value.Kind switch
        {
            CellValueKind.Number => FormulaValue.Scalar(value),
            CellValueKind.DateTime or CellValueKind.TimeSpan => FormulaValue.Number(value.ToSerialNumber(context.DateSystem)),
            CellValueKind.Boolean => FormulaValue.Number(value.GetBoolean() ? 1 : 0),
            CellValueKind.Error => FormulaValue.Scalar(value),
            _ => FormulaValue.Number(0),
        };
    }

    private static FormulaValue Type(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].IsArray)
            return FormulaValue.Number(64);

        CellValue value = Operators.Dereference(args[0], context);
        return FormulaValue.Number(value.Kind switch
        {
            CellValueKind.Text => 2,
            CellValueKind.Boolean => 4,
            CellValueKind.Error => 16,
            _ => 1,
        });
    }

    private static FormulaValue ErrorType(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue value = Operators.Dereference(args[0], context);
        if (!value.IsError)
            return FormulaValue.Error(ExcelError.NotAvailable);

        return FormulaValue.Number(value.GetError() switch
        {
            ExcelError.Null => 1,
            ExcelError.DivisionByZero => 2,
            ExcelError.Value => 3,
            ExcelError.Reference => 4,
            ExcelError.Name => 5,
            ExcelError.Number => 6,
            ExcelError.NotAvailable => 7,
            _ => 8,
        });
    }
}
