namespace Cellwright.Formulas.Functions;

/// <summary>
/// Implementation of a worksheet function. Arguments arrive in declaration order; the
/// evaluator has already turned references into values unless
/// <see cref="FunctionDefinition.WantsRawReferences"/> is set.
/// </summary>
internal delegate FormulaValue FunctionHandler(EvalContext context, ReadOnlySpan<FormulaValue> args);

/// <summary>
/// A function as seen by the parser and the evaluator: arity for validation and flags that
/// decide how arguments are prepared.
/// </summary>
internal sealed class FunctionDefinition
{
    public const int Unlimited = 255;

    public FunctionDefinition(string name, int minArgs, int maxArgs, FunctionHandler handler)
    {
        Name = name;
        MinArgs = minArgs;
        MaxArgs = maxArgs;
        Handler = handler;
    }

    /// <summary>Uppercase name as written in formulas.</summary>
    public string Name { get; }

    /// <summary>Smallest accepted argument count.</summary>
    public int MinArgs { get; }

    /// <summary>Largest accepted argument count, or <see cref="Unlimited"/>.</summary>
    public int MaxArgs { get; }

    /// <summary>The implementation.</summary>
    public FunctionHandler Handler { get; }

    /// <summary>Whether the function must be recalculated on every pass, e.g. <c>NOW</c>.</summary>
    public bool IsVolatile { get; init; }

    /// <summary>
    /// Whether references reach the handler untouched. Aggregates and functions that inspect
    /// geometry (<c>ROW</c>, <c>COLUMNS</c>, <c>OFFSET</c>) need this; scalar functions do not
    /// and get implicit intersection applied for them.
    /// </summary>
    public bool WantsRawReferences { get; init; }

    /// <summary>
    /// Whether the function decides itself what to do with error arguments. Normally an error
    /// argument short-circuits the call, which is wrong for <c>IFERROR</c> and the <c>IS…</c> family.
    /// </summary>
    public bool AcceptsErrors { get; init; }

    /// <summary>
    /// Whether the result depends on which rows are hidden. Only <c>SUBTOTAL</c> and
    /// <c>AGGREGATE</c> do, and hiding a row has to make exactly those recalculate even
    /// though no cell was written.
    /// </summary>
    public bool DependsOnVisibility { get; init; }
}
