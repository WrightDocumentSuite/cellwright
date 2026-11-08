using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Compile-time metadata of a <c>LAMBDA</c> expression inside a <see cref="FormulaProgram"/>.
/// </summary>
internal sealed class LambdaDefinition(int[] parameterSlots, int bodyStart)
{
    /// <summary>Local slot indices corresponding to the lambda's declared parameters.</summary>
    public int[] ParameterSlots { get; } = parameterSlots;

    /// <summary>Index of the first instruction of the lambda's body expression.</summary>
    public int BodyStart { get; } = bodyStart;

    /// <summary>Index of the closing <see cref="OpCode.EndLambda"/> instruction.</summary>
    public int BodyEnd { get; internal set; }
}

/// <summary>
/// Runtime closure created when a <c>LAMBDA</c> expression is evaluated.
/// </summary>
internal sealed class LambdaValue(
    FormulaProgram program,
    int lambdaIndex,
    FormulaValue[]? capturedLocals)
{
    public FormulaProgram Program { get; } = program;

    public int LambdaIndex { get; } = lambdaIndex;

    public FormulaValue[]? CapturedLocals { get; } = capturedLocals;

    public LambdaDefinition Definition => Program.Lambdas[LambdaIndex];

    public override string ToString() => "#CALC!";
}
