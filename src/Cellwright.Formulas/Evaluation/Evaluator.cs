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

    /// <summary>Evaluates the program in the context's current cell.</summary>
    public FormulaValue Evaluate(FormulaProgram program, EvalContext context)
    {
        _top = 0;
        EnsureCapacity(program.MaxStackDepth);
        if (_locals.Length < program.LocalNames.Length)
            _locals = new FormulaValue[program.LocalNames.Length];

        foreach (Instruction instruction in program.Instructions)
            Step(program, instruction, context);

        return _top == 1 ? _stack[0] : FormulaValue.Error(ExcelError.Value);
    }

    private void Step(FormulaProgram program, Instruction instruction, EvalContext context)
    {
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
                Push(FormulaValue.Blank);
                break;
            case OpCode.PushReference:
                PushReference(program.References[instruction.Operand], context);
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

            // The calculation of a LET is already on top of the stack; the marker only exists
            // so that rendering can put the bindings back in front of it.
            case OpCode.Let:
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

    private void PushReference(in FormulaRef reference, EvalContext context)
    {
        // A cell of another workbook has no worksheet behind it; its cached values arrive as
        // a plain value or array instead of a reference.
        if (reference.IsExternal)
        {
            Push(context.ResolveExternal(reference));
            return;
        }

        ReferenceValue? resolved = context.Resolve(reference);
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

        FormulaValue result = definition is null
            ? FormulaValue.Error(ExcelError.Name)
            : Invoke(definition, args, context);

        _top -= argCount;
        Push(result);
    }

    private static FormulaValue Invoke(FunctionDefinition definition, Span<FormulaValue> args, EvalContext context)
    {
        if (!definition.WantsRawReferences)
        {
            // Scalar functions see values, not geometry: references collapse through implicit
            // intersection before the handler runs.
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
