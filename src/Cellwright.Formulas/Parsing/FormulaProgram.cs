using Cellwright.Formulas.Functions;

namespace Cellwright.Formulas.Parsing;

/// <summary>
/// A parsed formula: a flat instruction array plus the operand pools it indexes. Programs are
/// immutable and shared between cells, so parsing the same text twice costs nothing after the
/// first time.
/// </summary>
internal sealed class FormulaProgram
{
    internal FormulaProgram(
        string text,
        Instruction[] instructions,
        double[] numbers,
        string[] strings,
        FormulaRef[] references,
        TableRef[] tableReferences,
        ArrayLiteral[] arrays,
        FunctionDefinition?[] functions,
        string[] functionNames,
        string[] localNames,
        LambdaDefinition[] lambdas,
        bool isVolatile,
        bool dependsOnVisibility)
    {
        LocalNames = localNames;
        Lambdas = lambdas;
        Text = text;
        Instructions = instructions;
        Numbers = numbers;
        Strings = strings;
        References = references;
        TableReferences = tableReferences;
        Arrays = arrays;
        Functions = functions;
        FunctionNames = functionNames;
        IsVolatile = isVolatile;
        DependsOnVisibility = dependsOnVisibility;
    }

    /// <summary>Source text the program was compiled from, without the leading <c>=</c>.</summary>
    public string Text { get; }

    internal Instruction[] Instructions { get; }

    internal double[] Numbers { get; }

    internal string[] Strings { get; }

    internal FormulaRef[] References { get; }

    /// <summary>Structured references, resolved against the workbook's tables rather than coordinates.</summary>
    internal TableRef[] TableReferences { get; }

    internal ArrayLiteral[] Arrays { get; }

    /// <summary>Resolved functions; a null entry is an unknown name that evaluates to <c>#NAME?</c>.</summary>
    internal FunctionDefinition?[] Functions { get; }

    /// <summary>Names as written, parallel to <see cref="Functions"/>, for faithful rendering.</summary>
    internal string[] FunctionNames { get; }

    /// <summary>Names of the <c>LET</c> and <c>LAMBDA</c> bindings, indexed by slot. Empty when the formula has none.</summary>
    internal string[] LocalNames { get; }

    /// <summary>Compiled <c>LAMBDA</c> descriptors inside this program.</summary>
    internal LambdaDefinition[] Lambdas { get; }

    /// <summary>Whether the formula must be recalculated on every pass.</summary>
    public bool IsVolatile { get; }

    /// <summary>Whether the result changes when rows are hidden or shown.</summary>
    public bool DependsOnVisibility { get; }

    /// <summary>Maximum depth the evaluation stack can reach.</summary>
    internal int MaxStackDepth
    {
        get
        {
            int depth = 0;
            int max = 0;
            foreach (Instruction instruction in Instructions)
            {
                depth += StackDelta(instruction);
                max = Math.Max(max, depth);
            }

            return max;
        }
    }

    private static int StackDelta(Instruction instruction) => instruction.Code switch
    {
        OpCode.Call => 1 - instruction.ArgCount,
        OpCode.Invoke => -instruction.ArgCount,
        OpCode.PushLambda => 0,
        OpCode.EndLambda => 0,
        <= OpCode.PushLocal => 1,
        OpCode.Negate or OpCode.Plus or OpCode.Percent or OpCode.Let => 0,
        _ => -1,
    };

    /// <inheritdoc/>
    public override string ToString() => Text;
}

/// <summary>An array constant such as <c>{1,2;3,4}</c>, stored row-major.</summary>
internal sealed class ArrayLiteral(int rows, int columns, CellValue[] values)
{
    public int Rows { get; } = rows;

    public int Columns { get; } = columns;

    public CellValue[] Values { get; } = values;
}
