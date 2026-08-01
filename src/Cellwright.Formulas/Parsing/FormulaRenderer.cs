using System.Text;

namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Turns a compiled program back into A1 formula text, optionally shifting relative
/// references. This is what expands shared formulas and what keeps formula text correct
/// after a sheet is renamed.
/// </summary>
internal static class FormulaRenderer
{
    /// <summary>
    /// Renders the program. <paramref name="rowDelta"/> and <paramref name="columnDelta"/>
    /// move relative references, as copying the formula would.
    /// </summary>
    public static string Render(FormulaProgram program, Workbook? workbook, int rowDelta = 0, int columnDelta = 0)
    {
        var stack = new List<Fragment>(8);
        List<(string Name, string Value)>? bindings = program.LocalNames.Length > 0 ? [] : null;
        foreach (Instruction instruction in program.Instructions)
            Apply(program, workbook, instruction, stack, bindings, rowDelta, columnDelta);

        return stack.Count == 1 ? stack[0].Text : program.Text;
    }

    private static void Apply(
        FormulaProgram program, Workbook? workbook, Instruction instruction,
        List<Fragment> stack, List<(string Name, string Value)>? bindings, int rowDelta, int columnDelta)
    {
        switch (instruction.Code)
        {
            case OpCode.PushLocal:
                Push(stack, program.LocalNames[instruction.Operand]);
                return;
            case OpCode.StoreLocal:
                bindings!.Add((program.LocalNames[instruction.Operand], Pop(stack).Text));
                return;
            case OpCode.Let:
                RenderLet(stack, bindings!, instruction.ArgCount);
                return;

            case OpCode.PushNumber:
                Push(stack, program.Numbers[instruction.Operand].ToString(System.Globalization.CultureInfo.InvariantCulture));
                break;
            case OpCode.PushString:
                Push(stack, $"\"{program.Strings[instruction.Operand].Replace("\"", "\"\"")}\"");
                break;
            case OpCode.PushBoolean:
                Push(stack, instruction.Operand != 0 ? "TRUE" : "FALSE");
                break;
            case OpCode.PushError:
                Push(stack, ((ExcelError)instruction.Operand).ToText());
                break;
            case OpCode.PushBlank:
                Push(stack, "");
                break;
            case OpCode.PushReference:
                Push(stack, RenderReference(program, workbook, program.References[instruction.Operand], rowDelta, columnDelta));
                break;
            case OpCode.PushName:
                Push(stack, program.Strings[instruction.Operand]);
                break;
            case OpCode.PushTableReference:
                // A structured reference names a table, not coordinates, so copying a formula
                // leaves it alone: the text it was written with is still the right text.
                Push(stack, program.TableReferences[instruction.Operand].Text);
                break;
            case OpCode.PushArray:
                Push(stack, RenderArray(program.Arrays[instruction.Operand]));
                break;
            case OpCode.Call:
                RenderCall(program, instruction, stack);
                break;
            case OpCode.Negate:
                RenderUnary(stack, "-", prefix: true);
                break;
            case OpCode.Plus:
                RenderUnary(stack, "+", prefix: true);
                break;
            case OpCode.Percent:
                RenderUnary(stack, "%", prefix: false);
                break;
            default:
                RenderBinary(stack, instruction.Code);
                break;
        }
    }

    /// <summary>
    /// Rebuilds a <c>LET</c> from the bindings stored while its body was rendered. The last
    /// <paramref name="count"/> of them belong to this one: an inner LET already took its own.
    /// </summary>
    private static void RenderLet(List<Fragment> stack, List<(string Name, string Value)> bindings, int count)
    {
        Fragment calculation = Pop(stack);
        int first = bindings.Count - count;
        var builder = new StringBuilder("LET(");
        for (int i = first; i < bindings.Count; i++)
            builder.Append(bindings[i].Name).Append(',').Append(bindings[i].Value).Append(',');

        bindings.RemoveRange(first, count);
        Push(stack, builder.Append(calculation.Text).Append(')').ToString());
    }

    private static void RenderCall(FormulaProgram program, Instruction instruction, List<Fragment> stack)
    {
        int argCount = instruction.ArgCount;
        var builder = new StringBuilder();
        builder.Append(program.FunctionNames[instruction.Operand]).Append('(');
        int first = stack.Count - argCount;
        for (int i = 0; i < argCount; i++)
        {
            if (i > 0)
                builder.Append(',');
            builder.Append(stack[first + i].Text);
        }

        stack.RemoveRange(first, argCount);
        builder.Append(')');
        Push(stack, builder.ToString());
    }

    private static void RenderUnary(List<Fragment> stack, string op, bool prefix)
    {
        Fragment operand = Pop(stack);
        string text = prefix
            ? op + Wrap(operand, RenderPrecedence.Unary)
            : Wrap(operand, RenderPrecedence.Percent) + op;
        Push(stack, text, prefix ? RenderPrecedence.Unary : RenderPrecedence.Percent);
    }

    private static void RenderBinary(List<Fragment> stack, OpCode code)
    {
        Fragment right = Pop(stack);
        Fragment left = Pop(stack);
        int precedence = RenderPrecedence.Of(code);
        string op = OperatorText(code);
        string text = code == OpCode.Union
            ? $"({left.Text},{right.Text})"
            : $"{Wrap(left, precedence)}{op}{Wrap(right, precedence + 1)}";
        Push(stack, text, code == OpCode.Union ? RenderPrecedence.Atom : precedence);
    }

    private static string RenderReference(
        FormulaProgram program, Workbook? workbook, FormulaRef reference, int rowDelta, int columnDelta)
    {
        FormulaRef shifted = reference.Shift(rowDelta, columnDelta);
        string prefix = RenderSheetPrefix(program, workbook, reference);
        if (!TryRenderPart(shifted.Start, out string start))
            return ExcelError.Reference.ToText();
        if (!shifted.IsRange)
            return shifted.IsSpilled ? prefix + start + "#" : prefix + start;
        if (!TryRenderPart(shifted.End, out string end))
            return ExcelError.Reference.ToText();
        return $"{prefix}{start}:{end}";
    }

    private static string RenderSheetPrefix(FormulaProgram program, Workbook? workbook, FormulaRef reference)
    {
        if (reference.SheetNameIndex < 0)
            return "";

        string name = ResolveName(program, workbook, reference.SheetId, reference.SheetNameIndex, reference.IsExternal);
        string? spanEnd = reference.SpanSheetNameIndex >= 0
            ? ResolveName(program, workbook, reference.SpanSheetId, reference.SpanSheetNameIndex, reference.IsExternal)
            : null;
        return FormulaText.QuotePrefix(reference.ExternalBook, name, spanEnd) + "!";
    }

    /// <summary>The sheet's current name when it still exists, the original spelling otherwise.</summary>
    private static string ResolveName(
        FormulaProgram program, Workbook? workbook, int sheetId, int nameIndex, bool external)
    {
        string original = program.Strings[nameIndex];
        return !external && sheetId >= 0
            ? workbook?.Sheets.FindById(sheetId)?.Name ?? original
            : original;
    }

    private static bool TryRenderPart(RefPart part, out string text)
    {
        text = "";
        switch (part.Kind)
        {
            case RefPartKind.Cell:
                if (!SheetLimits.IsValidRow(part.Row) || !SheetLimits.IsValidColumn(part.Column))
                    return false;
                text = $"{Anchor(part.ColumnAbsolute)}{ColumnName.Format(part.Column)}{Anchor(part.RowAbsolute)}{part.Row}";
                return true;

            case RefPartKind.Column:
                if (!SheetLimits.IsValidColumn(part.Column))
                    return false;
                text = Anchor(part.ColumnAbsolute) + ColumnName.Format(part.Column);
                return true;

            case RefPartKind.Row:
                if (!SheetLimits.IsValidRow(part.Row))
                    return false;
                text = Anchor(part.RowAbsolute) + part.Row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;

            default:
                return false;
        }
    }

    private static string RenderArray(ArrayLiteral array)
    {
        var builder = new StringBuilder("{");
        for (int row = 0; row < array.Rows; row++)
        {
            if (row > 0)
                builder.Append(';');
            for (int column = 0; column < array.Columns; column++)
            {
                if (column > 0)
                    builder.Append(',');
                builder.Append(RenderConstant(array.Values[(row * array.Columns) + column]));
            }
        }

        return builder.Append('}').ToString();
    }

    private static string RenderConstant(CellValue value) => value.Kind switch
    {
        CellValueKind.Text => $"\"{value.GetText().Replace("\"", "\"\"")}\"",
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        CellValueKind.Blank => "",
        _ => value.ToString(),
    };

    private static string Anchor(bool absolute) => absolute ? "$" : "";

    private static string OperatorText(OpCode code) => code switch
    {
        OpCode.Add => "+",
        OpCode.Subtract => "-",
        OpCode.Multiply => "*",
        OpCode.Divide => "/",
        OpCode.Power => "^",
        OpCode.Concat => "&",
        OpCode.Equal => "=",
        OpCode.NotEqual => "<>",
        OpCode.Less => "<",
        OpCode.LessOrEqual => "<=",
        OpCode.Greater => ">",
        OpCode.GreaterOrEqual => ">=",
        OpCode.Range => ":",
        OpCode.Intersect => " ",
        _ => throw new InvalidOperationException($"{code} has no operator text."),
    };

    private static string Wrap(Fragment fragment, int parentPrecedence) =>
        fragment.Precedence < parentPrecedence ? $"({fragment.Text})" : fragment.Text;

    private static void Push(List<Fragment> stack, string text, int precedence = RenderPrecedence.Atom) =>
        stack.Add(new Fragment(text, precedence));

    private static Fragment Pop(List<Fragment> stack)
    {
        if (stack.Count == 0)
            return new Fragment("", RenderPrecedence.Atom);
        Fragment fragment = stack[^1];
        stack.RemoveAt(stack.Count - 1);
        return fragment;
    }

    private readonly record struct Fragment(string Text, int Precedence);

    private static class RenderPrecedence
    {
        public const int Atom = 100;
        public const int Percent = 7;
        public const int Unary = 6;

        public static int Of(OpCode code) => code switch
        {
            OpCode.Equal or OpCode.NotEqual or OpCode.Less or
            OpCode.LessOrEqual or OpCode.Greater or OpCode.GreaterOrEqual => 1,
            OpCode.Concat => 2,
            OpCode.Add or OpCode.Subtract => 3,
            OpCode.Multiply or OpCode.Divide => 4,
            OpCode.Power => 5,
            OpCode.Range or OpCode.Intersect or OpCode.Union => 8,
            _ => Atom,
        };
    }
}
