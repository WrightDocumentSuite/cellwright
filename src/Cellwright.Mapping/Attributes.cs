namespace Cellwright.Mapping;

/// <summary>
/// Marks a partial type as an Excel row schema. The Cellwright.Mapping source generator
/// implements <see cref="IRowSchema{TRow}"/> on the type at compile time — no reflection,
/// Native AOT ready.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
public sealed class ExcelRowAttribute : Attribute
{
    /// <summary>
    /// Whether the sheet has a header row: written when producing files, used to bind
    /// columns by name when reading. Defaults to <see langword="true"/>.
    /// </summary>
    public bool HasHeader { get; set; } = true;
}

/// <summary>
/// Customizes how one property or positional record parameter maps to a column.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ExcelColumnAttribute : Attribute
{
    /// <summary>Maps with the property name as the header.</summary>
    public ExcelColumnAttribute()
    {
    }

    /// <summary>Maps with an explicit header text.</summary>
    public ExcelColumnAttribute(string header) => Header = header;

    /// <summary>Header text; the property name when omitted.</summary>
    public string? Header { get; }

    /// <summary>Explicit zero-based column position; declaration order when omitted.</summary>
    public int Order { get; set; } = -1;

    /// <summary>Number format code applied to the column when writing, e.g. <c>#,##0.00</c>.</summary>
    public string? Format { get; set; }
}

/// <summary>
/// Excludes a property or positional record parameter from mapping.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ExcelIgnoreAttribute : Attribute;
