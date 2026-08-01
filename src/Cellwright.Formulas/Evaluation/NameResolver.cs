using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Resolves defined names to values. A name refers to a formula, so resolving one means
/// evaluating it; nesting is bounded to keep mutually referring names from looping.
/// </summary>
internal static class NameResolver
{
    private const int MaxNameDepth = 8;

    /// <summary>Evaluates what a name refers to, or returns <c>#NAME?</c> when it is unknown.</summary>
    public static FormulaValue Resolve(string name, EvalContext context)
    {
        DefinedName? defined = context.Workbook.Names.Resolve(name, context.Sheet);
        if (defined is null)
            return FormulaValue.Error(ExcelError.Name);

        if (context.NameDepth >= MaxNameDepth)
            return FormulaValue.Error(ExcelError.Reference);

        FormulaProgram? program = context.GetProgram(defined.RefersTo);
        if (program is null)
            return FormulaValue.Error(ExcelError.Name);

        context.NameDepth++;
        try
        {
            return context.EvaluateNested(program);
        }
        finally
        {
            context.NameDepth--;
        }
    }
}
