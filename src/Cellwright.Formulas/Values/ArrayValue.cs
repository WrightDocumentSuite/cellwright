namespace Cellwright.Formulas;

/// <summary>
/// A rectangle of values held in memory: the result of an array constant or of reading a
/// range into a function that needs random access.
/// </summary>
/// <remarks>
/// Instances are immutable by contract. A materialized range is shared by every formula
/// reading it, so writing into <see cref="Values"/> would corrupt unrelated results; code
/// that transforms an array builds a new one instead.
/// </remarks>
internal sealed class ArrayValue
{
    /// <summary>Largest array the engine will materialize, guarding against runaway formulas.</summary>
    public const int MaxCells = 4_000_000;

    public ArrayValue(int rows, int columns, CellValue[] values)
    {
        Rows = rows;
        Columns = columns;
        Values = values;
    }

    public int Rows { get; }

    public int Columns { get; }

    public CellValue[] Values { get; }

    public int Count => Rows * Columns;

    public CellValue this[int row, int column] => Values[(row * Columns) + column];

    /// <summary>Wraps a single value as a one-by-one array.</summary>
    public static ArrayValue Single(CellValue value) => new(1, 1, [value]);

    /// <summary>
    /// Reads an element the way array arithmetic does: a single row or column is reused
    /// along its degenerate axis, matching Excel's broadcasting.
    /// </summary>
    public CellValue Broadcast(int row, int column)
    {
        int r = Rows == 1 ? 0 : row;
        int c = Columns == 1 ? 0 : column;
        if ((uint)r >= (uint)Rows || (uint)c >= (uint)Columns)
            return CellValue.Error(ExcelError.NotAvailable);
        return this[r, c];
    }

    /// <inheritdoc/>
    public override string ToString() => $"{{{Rows}x{Columns}}}";
}
