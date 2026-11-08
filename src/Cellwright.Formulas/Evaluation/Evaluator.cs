using Cellwright.Formulas.Functions;
using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Runs a compiled program on a reusable stack. One evaluator instance serves a whole
/// calculation pass, so after warm-up the only allocations are the results themselves.
/// </summary>
internal sealed class Evaluator
{
    private FormulaValue[] _stack = new FormulaValue[64];
    private FormulaValue[] _locals = [];
    private int _top;

    /// <summary>
    /// Evaluates the program in the context's current cell, optionally as a formula copied by
    /// the supplied row and column offsets.
    /// </summary>
    public FormulaValue Evaluate(
        FormulaProgram program, EvalContext context, int rowDelta = 0, int columnDelta = 0)
    {
        _top = 0;
        EnsureCapacity(program.MaxStackDepth + 8);
        if (_locals.Length < program.LocalNames.Length)
            _locals = new FormulaValue[program.LocalNames.Length];

        Instruction[] instructions = program.Instructions;
        for (int ip = 0; ip < instructions.Length; ip++)
            Step(program, ref ip, context, rowDelta, columnDelta);

        return _top == 1 ? _stack[0] : FormulaValue.Error(ExcelError.Value);
    }

    public FormulaValue EvaluateLambda(
        LambdaValue lambda, ReadOnlySpan<FormulaValue> args, EvalContext context)
    {
        FormulaProgram program = lambda.Program;
        _top = 0;
        EnsureCapacity(program.MaxStackDepth + 8);
        if (_locals.Length < program.LocalNames.Length)
            _locals = new FormulaValue[program.LocalNames.Length];

        if (lambda.CapturedLocals is { } captured)
            captured.CopyTo(_locals.AsSpan());

        int[] slots = lambda.Definition.ParameterSlots;
        for (int i = 0; i < slots.Length; i++)
            _locals[slots[i]] = i < args.Length ? args[i] : FormulaValue.Omitted;

        int end = lambda.Definition.BodyEnd - 1;
        for (int ip = lambda.Definition.BodyStart; ip < end; ip++)
            Step(program, ref ip, context, 0, 0);

        return _top == 1 ? _stack[0] : FormulaValue.Error(ExcelError.Value);
    }

    private void Step(
        FormulaProgram program,
        ref int ip,
        EvalContext context,
        int rowDelta,
        int columnDelta)
    {
        Instruction instruction = program.Instructions[ip];
        switch (instruction.Code)
        {
            case OpCode.PushNumber:
                Push(FormulaValue.Number(program.Numbers[instruction.Operand]));
                break;
            case OpCode.PushString:
                Push(FormulaValue.Text(program.Strings[instruction.Operand]));
                break;
            case OpCode.PushBoolean:
                Push(FormulaValue.Boolean(instruction.Operand != 0));
                break;
            case OpCode.PushError:
                Push(FormulaValue.Error((ExcelError)instruction.Operand));
                break;
            case OpCode.PushBlank:
                Push(FormulaValue.Omitted);
                break;
            case OpCode.PushReference:
                PushReference(program.References[instruction.Operand], context, rowDelta, columnDelta);
                break;
            case OpCode.PushName:
                Push(NameResolver.Resolve(program.Strings[instruction.Operand], context));
                break;
            case OpCode.PushTableReference:
                PushTableReference(program.TableReferences[instruction.Operand], context);
                break;
            case OpCode.PushArray:
                PushArray(program.Arrays[instruction.Operand]);
                break;
            case OpCode.PushLocal:
                Push(_locals[instruction.Operand]);
                break;
            case OpCode.StoreLocal:
                _locals[instruction.Operand] = Pop();
                break;
            case OpCode.Let or OpCode.EndLambda:
                break;
            case OpCode.PushLambda:
                PushLambda(program, instruction.Operand, ref ip);
                break;
            case OpCode.Invoke:
                InvokeOnStack(instruction.ArgCount, context);
                break;
            case OpCode.Call:
                Call(program, instruction, context);
                break;
            case OpCode.Negate or OpCode.Plus or OpCode.Percent:
                _stack[_top - 1] = Operators.Unary(instruction.Code, _stack[_top - 1], context);
                break;
            default:
                FormulaValue right = Pop();
                FormulaValue left = Pop();
                Push(Operators.Binary(instruction.Code, left, right, context));
                break;
        }
    }

    private void PushLambda(FormulaProgram program, int lambdaIndex, ref int ip)
    {
        FormulaValue[]? captured = program.LocalNames.Length > 0
            ? _locals.AsSpan(0, program.LocalNames.Length).ToArray()
            : null;
        Push(FormulaValue.Lambda(new LambdaValue(program, lambdaIndex, captured)));
        ip = program.Lambdas[lambdaIndex].BodyEnd - 1;
    }

    private void InvokeOnStack(int argCount, EvalContext context)
    {
        Span<FormulaValue> args = _stack.AsSpan(_top - argCount, argCount);
        FormulaValue target = _stack[_top - argCount - 1];
        FormulaValue result = target.IsLambda
            ? context.InvokeLambda(target.AsLambda, args)
            : target.IsError ? target : FormulaValue.Error(ExcelError.Value);

        _top -= argCount + 1;
        Push(result);
    }

    private void PushReference(
        in FormulaRef reference, EvalContext context, int rowDelta, int columnDelta)
    {
        FormulaRef shifted = rowDelta == 0 && columnDelta == 0
            ? reference
            : reference.Shift(rowDelta, columnDelta);

        if (shifted.IsExternal)
        {
            Push(context.ResolveExternal(shifted));
            return;
        }

        ReferenceValue? resolved = context.Resolve(shifted);
        Push(resolved is null ? FormulaValue.Error(ExcelError.Reference) : FormulaValue.Reference(resolved));
    }

    private void PushTableReference(in TableRef reference, EvalContext context)
    {
        ReferenceValue? resolved = context.Resolve(reference);
        Push(resolved is null ? FormulaValue.Error(ExcelError.Reference) : FormulaValue.Reference(resolved));
    }

    private void PushArray(ArrayLiteral literal) =>
        Push(FormulaValue.Array(new ArrayValue(literal.Rows, literal.Columns, literal.Values)));

    private void Call(FormulaProgram program, Instruction instruction, EvalContext context)
    {
        int argCount = instruction.ArgCount;
        FunctionDefinition? definition = program.Functions[instruction.Operand];
        Span<FormulaValue> args = _stack.AsSpan(_top - argCount, argCount);

        FormulaValue result;
        if (definition is not null)
        {
            result = Invoke(definition, args, context);
        }
        else
        {
            FormulaValue named = NameResolver.Resolve(program.FunctionNames[instruction.Operand], context);
            result = named.IsLambda
                ? context.InvokeLambda(named.AsLambda, args)
                : FormulaValue.Error(ExcelError.Name);
        }

        _top -= argCount;
        Push(result);
    }

    private static FormulaValue Invoke(FunctionDefinition definition, Span<FormulaValue> args, EvalContext context)
    {
        if (!definition.WantsRawReferences)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Kind == FormulaValueKind.Reference)
                    args[i] = FormulaValue.Scalar(Operators.Dereference(args[i], context));
            }
        }

        if (!definition.AcceptsErrors)
        {
            foreach (ref readonly FormulaValue arg in args)
            {
                if (arg.IsError)
                    return arg;
            }
        }

        context.CancellationToken.ThrowIfCancellationRequested();
        return definition.Handler(context, args);
    }

    private void Push(FormulaValue value)
    {
        EnsureCapacity(_top + 1);
        _stack[_top++] = value;
    }

    private FormulaValue Pop() => _top > 0 ? _stack[--_top] : FormulaValue.Error(ExcelError.Value);

    private void EnsureCapacity(int required)
    {
        if (_stack.Length < required)
            Array.Resize(ref _stack, Math.Max(required, _stack.Length * 2));
    }
}
