namespace Cellwright;

/// <summary>
/// A recoverable fidelity problem encountered while reading, editing, copying, or saving a
/// workbook. The library takes the documented safe path and keeps going instead of hiding the
/// limitation or failing the whole operation.
/// </summary>
/// <param name="Code">Stable machine-readable code, e.g. <c>BadNumber</c>.</param>
/// <param name="Message">Human-readable description.</param>
/// <param name="Cell">The affected cell, when known.</param>
public sealed record WorkbookWarning(string Code, string Message, CellRef? Cell = null)
{
    /// <inheritdoc/>
    public override string ToString() => Cell is { } cell ? $"[{Code}] {cell}: {Message}" : $"[{Code}] {Message}";
}
