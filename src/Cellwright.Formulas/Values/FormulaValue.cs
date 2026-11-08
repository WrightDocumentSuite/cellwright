namespace Cellwright.Formulas;

/// <summary>What a <see cref="FormulaValue"/> carries.</summary>
public enum FormulaValueKind : byte
{
    /// <summary>A single value: number, text, logical, error or blank.</summary>
    Scalar,

    /// <summary>A reference to one or more worksheet areas.</summary>
    Reference,

    /// <summary>An in-memory rectangle of values.</summary>
    Array,

    /// <summary>A callable <c>LAMBDA</c> closure.</summary>
    Lambda,
}

/// <summary>
/// A value flowing through the evaluator: a scalar, a worksheet reference, an array, or a lambda.
/// Scalars reuse <see cref="CellValue"/>, so numbers and dates cost no allocation.
/// </summary>
public readonly struct FormulaValue
{
    private static readonly object OmittedSentinel = new();

    private readonly CellValue _scalar;
    private readonly object? _payload;

    private FormulaValue(CellValue scalar, object? payload = null)
    {
        _scalar = scalar;
        _payload = payload;
        Kind = FormulaValueKind.Scalar;
    }

    private FormulaValue(object payload, FormulaValueKind kind)
    {
        _scalar = default;
        _payload = payload;
        Kind = kind;
    }

    /// <summary>What the value carries.</summary>
    public FormulaValueKind Kind { get; }

    /// <summary>The blank value.</summary>
    public static FormulaValue Blank => default;

    /// <summary>An omitted argument value (behaves as blank, but <c>ISOMITTED</c> returns <see langword="true"/>).</summary>
    internal static FormulaValue Omitted => new(CellValue.Blank, OmittedSentinel);

    /// <summary>Wraps a scalar.</summary>
    public static FormulaValue Scalar(CellValue value) => new(value);

    /// <summary>Wraps a number.</summary>
    public static FormulaValue Number(double value) => new(CellValue.Number(value));

    /// <summary>Wraps text.</summary>
    public static FormulaValue Text(string value) => new(CellValue.Text(value));

    /// <summary>Wraps a logical value.</summary>
    public static FormulaValue Boolean(bool value) => new(CellValue.Boolean(value));

    /// <summary>Wraps an error.</summary>
    public static FormulaValue Error(ExcelError error) => new(CellValue.Error(error));

    internal static FormulaValue Reference(ReferenceValue reference) => new(reference, FormulaValueKind.Reference);

    internal static FormulaValue Array(ArrayValue array) => new(array, FormulaValueKind.Array);

    internal static FormulaValue Lambda(LambdaValue lambda) => new(lambda, FormulaValueKind.Lambda);

    /// <summary>Whether the value is a scalar of any kind.</summary>
    public bool IsScalar => Kind == FormulaValueKind.Scalar;

    /// <summary>Whether the value is a worksheet reference.</summary>
    public bool IsReference => Kind == FormulaValueKind.Reference;

    /// <summary>Whether the value is an array.</summary>
    public bool IsArray => Kind == FormulaValueKind.Array;

    /// <summary>Whether the value is a callable <c>LAMBDA</c>.</summary>
    public bool IsLambda => Kind == FormulaValueKind.Lambda;

    /// <summary>Whether the value was produced by an omitted argument in a call.</summary>
    internal bool IsOmitted => ReferenceEquals(_payload, OmittedSentinel);

    /// <summary>Whether the value is a scalar error.</summary>
    public bool IsError => Kind == FormulaValueKind.Scalar && _scalar.IsError;

    /// <summary>Whether the value is a scalar blank.</summary>
    public bool IsBlank => Kind == FormulaValueKind.Scalar && _scalar.IsBlank;

    /// <summary>The scalar payload; blank for references and arrays.</summary>
    public CellValue AsScalar => _scalar;

    internal ReferenceValue AsReference => (ReferenceValue)_payload!;

    internal ArrayValue AsArray => (ArrayValue)_payload!;

    internal LambdaValue AsLambda => (LambdaValue)_payload!;

    /// <summary>Wraps a cell value.</summary>
    public static implicit operator FormulaValue(CellValue value) => new(value);

    /// <inheritdoc/>
    public override string ToString() => Kind switch
    {
        FormulaValueKind.Reference => AsReference.ToString(),
        FormulaValueKind.Array => AsArray.ToString(),
        FormulaValueKind.Lambda => AsLambda.ToString(),
        _ => _scalar.ToString(),
    };
}
