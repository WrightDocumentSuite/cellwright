namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Instructions of a compiled formula. The program is reverse-Polish: operands push,
/// operators pop.
/// </summary>
internal enum OpCode : byte
{
    /// <summary>Pushes a number from the number pool.</summary>
    PushNumber,

    /// <summary>Pushes text from the string pool.</summary>
    PushString,

    /// <summary>Pushes a logical value; the operand is 0 or 1.</summary>
    PushBoolean,

    /// <summary>Pushes an error value; the operand is an <see cref="ExcelError"/>.</summary>
    PushError,

    /// <summary>Pushes a blank, produced by an omitted argument such as the middle of <c>IF(a,,b)</c>.</summary>
    PushBlank,

    /// <summary>Pushes a reference from the reference pool.</summary>
    PushReference,

    /// <summary>Pushes the value of a defined name from the name pool.</summary>
    PushName,

    /// <summary>Pushes an array constant from the array pool.</summary>
    PushArray,

    /// <summary>Pushes a structured table reference, resolved against the workbook's tables.</summary>
    PushTableReference,

    /// <summary>Pushes the value bound to a <c>LET</c> slot.</summary>
    PushLocal,

    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concat,

    Equal,
    NotEqual,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,

    /// <summary>Unary minus.</summary>
    Negate,

    /// <summary>Unary plus: a no-op that still forces numeric coercion.</summary>
    Plus,

    /// <summary>Postfix percent: divides by 100.</summary>
    Percent,

    /// <summary>The <c>:</c> operator joining two references into a range.</summary>
    Range,

    /// <summary>The <c>,</c> operator joining two references into a multi-area reference.</summary>
    Union,

    /// <summary>The space operator intersecting two references.</summary>
    Intersect,

    /// <summary>Calls a function; the operand indexes the function pool, ArgCount gives arity.</summary>
    Call,

    /// <summary>Pops a value into a <c>LET</c> slot named by the operand.</summary>
    StoreLocal,

    /// <summary>
    /// Closes a <c>LET</c>. Evaluation ignores it — the calculation is already on the stack —
    /// but rendering needs it to know how many bindings to put back in front.
    /// </summary>
    Let,

    /// <summary>Pushes a <c>LAMBDA</c> closure from the lambda pool and jumps past its body.</summary>
    PushLambda,

    /// <summary>Marks the end of a <c>LAMBDA</c> body for formula rendering.</summary>
    EndLambda,

    /// <summary>Invokes a <c>LAMBDA</c> closure on the stack with <c>ArgCount</c> arguments.</summary>
    Invoke,
}

/// <summary>
/// One instruction: an opcode, an argument count for calls and an index into an operand pool.
/// Kept at eight bytes so a program is a compact, cache-friendly array.
/// </summary>
internal readonly record struct Instruction(OpCode Code, byte ArgCount, int Operand)
{
    public Instruction(OpCode code, int operand) : this(code, 0, operand)
    {
    }

    public Instruction(OpCode code) : this(code, 0, 0)
    {
    }
}
