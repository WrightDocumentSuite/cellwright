namespace Cellwright.Formulas;

/// <summary>
/// Walks the cells of a reference in row-major order across every area of a union, reading
/// straight from the worksheet's column slices. No intermediate collection is built, so a
/// <c>SUM</c> over a large range allocates nothing.
/// </summary>
internal ref struct ReferenceEnumerator
{
    private readonly EvalContext _context;
    private ReferenceValue? _area;
    private int _row;
    private int _column;

    public ReferenceEnumerator(ReferenceValue reference, EvalContext context)
    {
        _context = context;
        _area = reference;
        _row = reference.FirstRow;
        _column = reference.FirstColumn - 1;
    }

    /// <summary>Value of the cell the enumerator stands on.</summary>
    public CellValue Current { get; private set; }

    /// <summary>Position of the cell the enumerator stands on.</summary>
    public CellRef Position => new(_row, _column);

    /// <summary>Sheet of the cell the enumerator stands on.</summary>
    public readonly Worksheet Sheet => _area!.Sheet;

    public readonly ReferenceEnumerator GetEnumerator() => this;

    public bool MoveNext()
    {
        while (_area is not null)
        {
            _column++;
            if (_column > _area.LastColumn)
            {
                _column = _area.FirstColumn;
                _row++;
            }

            if (_row > _area.LastRow)
            {
                _area = _area.Next;
                if (_area is null)
                    return false;
                _row = _area.FirstRow;
                _column = _area.FirstColumn;
            }

            Current = _context.ReadCell(_area.Sheet, _row, _column);
            return true;
        }

        return false;
    }
}

/// <summary>
/// Helpers for turning references and arrays into the shapes functions expect.
/// </summary>
internal static class RangeReader
{
    /// <summary>Total number of cells a reference covers, across all areas of a union.</summary>
    public static long CountCells(ReferenceValue reference)
    {
        long total = 0;
        for (ReferenceValue? area = reference; area is not null; area = area.Next)
            total += (long)area.RowCount * area.ColumnCount;
        return total;
    }

    /// <summary>Materializes a reference into an array, guarding against runaway sizes.</summary>
    public static ArrayValue? ToArray(ReferenceValue reference, EvalContext context) =>
        Materialize(reference, context)?.Array;

    /// <summary>
    /// Reads a reference into a rectangle, reusing the cached copy when the same range was
    /// already read and nothing has been written into it since. The returned handle also
    /// carries the lookup indexes, so callers doing exact matches can avoid scanning.
    /// </summary>
    /// <remarks>
    /// Only single-area references are cached: a union materializes as its bounding box, and
    /// keying that by corners would let two different unions share one entry.
    /// </remarks>
    public static CachedRange? Materialize(ReferenceValue reference, EvalContext context)
    {
        ReferenceValue bounds = reference.Bounds();
        long cells = (long)bounds.RowCount * bounds.ColumnCount;
        if (cells > ArrayValue.MaxCells)
            return null;

        bool cacheable = reference.Next is null && cells >= RangeCache.MinCells;
        var key = new RangeKey(bounds.Sheet.Id, bounds.FirstRow, bounds.FirstColumn, bounds.LastRow, bounds.LastColumn);
        if (cacheable && context.Ranges.TryGet(key, out CachedRange hit))
            return hit;

        var values = new CellValue[cells];
        int index = 0;
        for (int row = bounds.FirstRow; row <= bounds.LastRow; row++)
        {
            for (int column = bounds.FirstColumn; column <= bounds.LastColumn; column++)
                values[index++] = context.ReadCell(bounds.Sheet, row, column);
        }

        var materialized = new CachedRange(key, new ArrayValue(bounds.RowCount, bounds.ColumnCount, values));
        if (cacheable)
            context.Ranges.Add(materialized);

        return materialized;
    }

    /// <summary>
    /// Applies implicit intersection: a reference used where a single value is expected is
    /// narrowed to the row or column of the formula's own cell.
    /// </summary>
    public static CellValue Intersect(ReferenceValue reference, EvalContext context)
    {
        if (reference.IsSingleCell)
            return context.ReadCell(reference.Sheet, reference.FirstRow, reference.FirstColumn);

        if (reference.Next is not null)
            return CellValue.Error(ExcelError.Value);

        CellRef cell = context.Cell;
        if (reference.RowCount == 1 && cell.Column >= reference.FirstColumn && cell.Column <= reference.LastColumn)
            return context.ReadCell(reference.Sheet, reference.FirstRow, cell.Column);

        if (reference.ColumnCount == 1 && cell.Row >= reference.FirstRow && cell.Row <= reference.LastRow)
            return context.ReadCell(reference.Sheet, cell.Row, reference.FirstColumn);

        return CellValue.Error(ExcelError.Value);
    }
}
