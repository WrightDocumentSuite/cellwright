namespace Cellwright.Mapping;

/// <summary>
/// Compile-time row schema implemented by the source generator on types marked with
/// <see cref="ExcelRowAttribute"/>. Static abstract members keep mapping stateless,
/// reflection-free and Native AOT compatible.
/// </summary>
/// <typeparam name="TRow">The row type itself (curiously recurring).</typeparam>
public interface IRowSchema<TRow> where TRow : IRowSchema<TRow>
{
    /// <summary>Number of mapped columns.</summary>
    static abstract int ColumnCount { get; }

    /// <summary>Whether files carry a header row for this schema.</summary>
    static abstract bool HasHeader { get; }

    /// <summary>Header texts in column order.</summary>
    static abstract ReadOnlySpan<string> Headers { get; }

    /// <summary>
    /// Materializes a row. <paramref name="columnMap"/> maps schema positions to zero-based
    /// sheet columns (a negative entry means the column is absent and reads blank).
    /// </summary>
    static abstract TRow Read(scoped ReadOnlySpan<CellValue> cells, scoped ReadOnlySpan<int> columnMap);

    /// <summary>Serializes a row into a buffer of <see cref="IRowSchema{TRow}.ColumnCount"/> cells.</summary>
    static abstract void Write(TRow row, scoped Span<CellValue> cells);

    /// <summary>
    /// Registers the per-column styles (number formats) declared via
    /// <see cref="ExcelColumnAttribute.Format"/> and fills the parallel style buffer.
    /// </summary>
    static abstract void RegisterStyles(Func<CellStyle, StyleId> registerStyle, scoped Span<StyleId> styles);
}
