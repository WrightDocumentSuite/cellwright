using Cellwright.Formulas.Functions;

namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Accumulates instructions and operand pools while parsing. Operands are deduplicated, so a
/// formula that mentions the same constant or reference twice stores it once.
/// </summary>
internal sealed class FormulaProgramBuilder
{
    private readonly List<Instruction> _instructions = [];
    private readonly List<double> _numbers = [];
    private readonly List<string> _strings = [];
    private readonly List<FormulaRef> _references = [];
    private readonly List<TableRef> _tableReferences = [];
    private readonly List<ArrayLiteral> _arrays = [];
    private readonly List<FunctionDefinition?> _functions = [];
    private readonly List<string> _functionNames = [];
    private readonly List<string> _localNames = [];
    private readonly List<LambdaDefinition> _lambdas = [];
    private readonly Dictionary<string, int> _stringIndex = new(StringComparer.Ordinal);

    public bool IsVolatile { get; private set; }

    public bool DependsOnVisibility { get; private set; }

    public void Emit(OpCode code) => _instructions.Add(new Instruction(code));

    public void Emit(OpCode code, int operand) => _instructions.Add(new Instruction(code, operand));

    /// <summary>Reserves a slot for a <c>LET</c> binding and returns its index.</summary>
    public int DeclareLocal(string name)
    {
        _localNames.Add(name);
        return _localNames.Count - 1;
    }

    /// <summary>Emits the marker that closes a <c>LET</c> with the given number of bindings.</summary>
    public void EmitLet(int bindings) =>
        _instructions.Add(new Instruction(OpCode.Let, checked((byte)bindings), 0));

    /// <summary>Starts a <c>LAMBDA</c> definition and emits <see cref="OpCode.PushLambda"/>.</summary>
    public int BeginLambda(int[] parameterSlots)
    {
        int index = _lambdas.Count;
        _instructions.Add(new Instruction(OpCode.PushLambda, index));
        _lambdas.Add(new LambdaDefinition(parameterSlots, _instructions.Count));
        return index;
    }

    /// <summary>Closes a <c>LAMBDA</c> definition and records its end instruction index.</summary>
    public void EndLambda(int lambdaIndex)
    {
        _instructions.Add(new Instruction(OpCode.EndLambda, lambdaIndex));
        _lambdas[lambdaIndex].BodyEnd = _instructions.Count;
    }

    /// <summary>Emits an invocation of a callable value on the stack.</summary>
    public void EmitInvoke(int argCount) =>
        _instructions.Add(new Instruction(OpCode.Invoke, checked((byte)argCount), 0));

    public void EmitNumber(double value)
    {
        int index = _numbers.IndexOf(value);
        if (index < 0)
        {
            index = _numbers.Count;
            _numbers.Add(value);
        }

        _instructions.Add(new Instruction(OpCode.PushNumber, index));
    }

    public void EmitString(string value) => _instructions.Add(new Instruction(OpCode.PushString, InternString(value)));

    public void EmitBoolean(bool value) => _instructions.Add(new Instruction(OpCode.PushBoolean, value ? 1 : 0));

    public void EmitError(ExcelError error) => _instructions.Add(new Instruction(OpCode.PushError, (int)error));

    public void EmitReference(FormulaRef reference)
    {
        int index = _references.IndexOf(reference);
        if (index < 0)
        {
            index = _references.Count;
            _references.Add(reference);
        }

        _instructions.Add(new Instruction(OpCode.PushReference, index));
    }

    public void EmitTableReference(TableRef reference)
    {
        int index = _tableReferences.IndexOf(reference);
        if (index < 0)
        {
            index = _tableReferences.Count;
            _tableReferences.Add(reference);
        }

        _instructions.Add(new Instruction(OpCode.PushTableReference, index));
    }

    public void EmitName(string name) => _instructions.Add(new Instruction(OpCode.PushName, InternString(name)));

    public void EmitArray(ArrayLiteral array)
    {
        _arrays.Add(array);
        _instructions.Add(new Instruction(OpCode.PushArray, _arrays.Count - 1));
    }

    /// <summary>Emits a call; <paramref name="definition"/> is null for an unknown function.</summary>
    public void EmitCall(FunctionDefinition? definition, string name, int argCount)
    {
        _functions.Add(definition);
        _functionNames.Add(definition?.Name ?? name);
        if (definition?.IsVolatile == true)
            IsVolatile = true;
        if (definition?.DependsOnVisibility == true)
            DependsOnVisibility = true;
        _instructions.Add(new Instruction(OpCode.Call, checked((byte)argCount), _functions.Count - 1));
    }

    /// <summary>Interns a string and returns its pool index.</summary>
    public int InternString(string value)
    {
        if (_stringIndex.TryGetValue(value, out int index))
            return index;

        index = _strings.Count;
        _strings.Add(value);
        _stringIndex[value] = index;
        return index;
    }

    public FormulaProgram Build(string text) => new(
        text,
        [.. _instructions],
        [.. _numbers],
        [.. _strings],
        [.. _references],
        [.. _tableReferences],
        [.. _arrays],
        [.. _functions],
        [.. _functionNames],
        [.. _localNames],
        [.. _lambdas],
        IsVolatile,
        DependsOnVisibility);
}
