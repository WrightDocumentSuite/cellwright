using System.Collections;

namespace Cellwright;

/// <summary>
/// Lightweight facade over a rectangular range of a <see cref="Worksheet"/>.
/// Bulk operations iterate the underlying slices directly. (Named <c>CellRange</c>
/// to avoid colliding with <see cref="System.Range"/>.)
/// </summary>
public readonly struct CellRange : IEnumerable<Cell>
{
    private readonly Worksheet _sheet;

    internal CellRange(Worksheet sheet, RangeRef reference)
    {
        _sheet = sheet;
        Ref = reference;
    }

    internal Worksheet Sheet => _sheet;

    /// <summary>The range's reference.</summary>
    public RangeRef Ref { get; }

    /// <summary>Writes one value into every cell of the range.</summary>
    public CellRange SetValue(CellValue value)
    {
        foreach (Cell cell in this)
            cell.SetValue(value);
        return this;
    }

    /// <summary>Applies a registered style handle to every cell.</summary>
    public CellRange SetStyle(StyleId style)
    {
        foreach (Cell cell in this)
            cell.SetStyle(style);
        return this;
    }

    /// <summary>Registers and applies a style record to every cell.</summary>
    public CellRange SetStyle(CellStyle style) => SetStyle(_sheet.Workbook.RegisterStyle(style));

    /// <summary>Clears values, formulas and styles of all cells.</summary>
    public CellRange Clear()
    {
        foreach (Cell cell in this)
            cell.Clear();
        return this;
    }

    /// <summary>Merges the range into one display cell.</summary>
    public CellRange Merge()
    {
        _sheet.Merge(Ref);
        return this;
    }

    /// <summary>Enumerates cells row-major.</summary>
    public IEnumerator<Cell> GetEnumerator()
    {
        for (int row = Ref.Start.Row; row <= Ref.End.Row; row++)
        {
            for (int column = Ref.Start.Column; column <= Ref.End.Column; column++)
                yield return new Cell(_sheet, new CellRef(row, column));
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
