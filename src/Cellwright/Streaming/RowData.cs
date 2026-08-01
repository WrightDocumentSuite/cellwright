namespace Cellwright;

/// <summary>
/// One worksheet row produced by <see cref="XlsxReader"/>. The struct is a view over a buffer
/// that the reader reuses for the next row: consume (or copy via <see cref="ToArray"/>) before
/// advancing the enumeration.
/// </summary>
public readonly struct RowData
{
    private readonly CellValue[] _cells;
    private readonly int _count;

    internal RowData(CellValue[] cells, int count, int rowNumber)
    {
        _cells = cells;
        _count = count;
        RowNumber = rowNumber;
    }

    /// <summary>The 1-based row number. Empty rows are skipped, so numbers may jump.</summary>
    public int RowNumber { get; }

    /// <summary>Number of cells up to the last occupied column; gaps are <see cref="CellValue.Blank"/>.</summary>
    public int Count => _count;

    /// <summary>The cells of the row; index 0 is column A.</summary>
    public ReadOnlySpan<CellValue> Cells => _cells.AsSpan(0, _count);

    /// <summary>Returns the cell at a zero-based index, or blank when the row is shorter.</summary>
    public CellValue GetValueOrBlank(int index) =>
        (uint)index < (uint)_count ? _cells[index] : CellValue.Blank;

    /// <summary>Copies the row out of the shared buffer.</summary>
    public CellValue[] ToArray() => Cells.ToArray();
}
