namespace Cellwright;

/// <summary>
/// A recoverable problem encountered while reading a workbook. The library substitutes
/// a safe value and keeps going instead of failing the whole load.
/// </summary>
/// <param name="Code">Stable machine-readable code, e.g. <c>BadNumber</c>.</param>
/// <param name="Message">Human-readable description.</param>
/// <param name="Cell">The affected cell, when known.</param>
public sealed record WorkbookWarning(string Code, string Message, CellRef? Cell = null)
{
    /// <inheritdoc/>
    public override string ToString() => Cell is { } cell ? $"[{Code}] {cell}: {Message}" : $"[{Code}] {Message}";
}
