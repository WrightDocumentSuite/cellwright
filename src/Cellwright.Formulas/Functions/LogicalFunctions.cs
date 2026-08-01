namespace Cellwright.Formulas.Functions;

/// <summary>
/// Conditional and boolean functions. Arguments are evaluated before the call, so these
/// choose between already-computed values rather than short-circuiting; the result matches
/// Excel because worksheet functions are free of side effects.
/// </summary>
internal static class LogicalFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("IF", 2, 3, If, acceptsErrors: true);
        builder.Add("IFS", 2, FunctionDefinition.Unlimited, Ifs, acceptsErrors: true);
        builder.Add("IFERROR", 2, 2, IfError, acceptsErrors: true);
        builder.Add("IFNA", 2, 2, IfNa, acceptsErrors: true);
        builder.Add("AND", 1, FunctionDefinition.Unlimited, (c, a) => Fold(c, a, all: true), wantsRawReferences: true);
        builder.Add("OR", 1, FunctionDefinition.Unlimited, (c, a) => Fold(c, a, all: false), wantsRawReferences: true);
        builder.Add("XOR", 1, FunctionDefinition.Unlimited, Xor, wantsRawReferences: true);
        builder.Add("NOT", 1, 1, Not);
        builder.Add("TRUE", 0, 0, static (_, _) => FormulaValue.Boolean(true));
        builder.Add("FALSE", 0, 0, static (_, _) => FormulaValue.Boolean(false));
        builder.Add("SWITCH", 3, FunctionDefinition.Unlimited, Switch, acceptsErrors: true);
    }

    private static FormulaValue If(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args[0].IsError)
            return args[0];
        if (!Args.TryBoolean(context, args[0], out bool condition, out CellValue error))
            return FormulaValue.Scalar(error);

        if (condition)
            return args[1];
        return args.Length > 2 ? args[2] : FormulaValue.Boolean(false);
    }

    private static FormulaValue Ifs(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        if (args.Length % 2 != 0)
            return FormulaValue.Error(ExcelError.NotAvailable);

        for (int i = 0; i < args.Length; i += 2)
        {
            if (args[i].IsError)
                return args[i];
            if (!Args.TryBoolean(context, args[i], out bool condition, out CellValue error))
                return FormulaValue.Scalar(error);
            if (condition)
                return args[i + 1];
        }

        return FormulaValue.Error(ExcelError.NotAvailable);
    }

    private static FormulaValue IfError(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        IsErrorValue(context, args[0]) ? args[1] : args[0];

    private static FormulaValue IfNa(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue value = Operators.Dereference(args[0], context);
        return value.IsError && value.GetError() == ExcelError.NotAvailable ? args[1] : args[0];
    }

    private static bool IsErrorValue(EvalContext context, in FormulaValue value) =>
        Operators.Dereference(value, context).IsError;

    private static FormulaValue Fold(EvalContext context, ReadOnlySpan<FormulaValue> args, bool all)
    {
        bool result = all;
        bool sawValue = false;
        foreach (ref readonly FormulaValue arg in args)
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!TryLogical(value, fromRange, out bool logical, out CellValue error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                sawValue = true;
                result = all ? result && logical : result || logical;
            }
        }

        return sawValue ? FormulaValue.Boolean(result) : FormulaValue.Error(ExcelError.Value);
    }

    private static FormulaValue Xor(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        int trueCount = 0;
        bool sawValue = false;
        foreach (ref readonly FormulaValue arg in args)
        {
            bool fromRange = arg.Kind != FormulaValueKind.Scalar;
            foreach (CellValue value in new ArgumentValues(arg, context))
            {
                if (!TryLogical(value, fromRange, out bool logical, out CellValue error))
                {
                    if (error.IsError)
                        return FormulaValue.Scalar(error);
                    continue;
                }

                sawValue = true;
                if (logical)
                    trueCount++;
            }
        }

        return sawValue ? FormulaValue.Boolean(trueCount % 2 == 1) : FormulaValue.Error(ExcelError.Value);
    }

    private static FormulaValue Not(EvalContext context, ReadOnlySpan<FormulaValue> args) =>
        Args.TryBoolean(context, args[0], out bool value, out CellValue error)
            ? FormulaValue.Boolean(!value)
            : FormulaValue.Scalar(error);

    private static FormulaValue Switch(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        CellValue subject = Operators.Dereference(args[0], context);
        int pairs = (args.Length - 1) / 2;
        for (int i = 0; i < pairs; i++)
        {
            CellValue candidate = Operators.Dereference(args[1 + (i * 2)], context);
            if (Coercion.Compare(subject, candidate, context.DateSystem) == 0)
                return args[2 + (i * 2)];
        }

        // An odd trailing argument is the default.
        return args.Length % 2 == 0 ? args[^1] : FormulaValue.Error(ExcelError.NotAvailable);
    }

    /// <summary>Logical coercion for AND/OR: text and blanks inside ranges are ignored.</summary>
    private static bool TryLogical(in CellValue value, bool fromRange, out bool logical, out CellValue error)
    {
        logical = false;
        error = CellValue.Blank;
        switch (value.Kind)
        {
            case CellValueKind.Boolean:
                logical = value.GetBoolean();
                return true;
            case CellValueKind.Number:
                logical = value.GetNumber() != 0;
                return true;
            case CellValueKind.Error:
                error = value;
                return false;
            case CellValueKind.Blank:
                return false;
            default:
                if (fromRange)
                    return false;
                CellValue converted = Coercion.ToBoolean(value);
                if (converted.IsError)
                {
                    error = converted;
                    return false;
                }

                logical = converted.GetBoolean();
                return true;
        }
    }
}
