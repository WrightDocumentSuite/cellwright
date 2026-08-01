namespace Cellwright.Formulas;

/// <summary>
/// One or more rectangular worksheet areas. Unions chain through <see cref="Next"/>, which
/// keeps the common single-area case free of collections.
/// </summary>
internal sealed class ReferenceValue
{
    public ReferenceValue(Worksheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        Sheet = sheet;
        FirstRow = Math.Min(firstRow, lastRow);
        FirstColumn = Math.Min(firstColumn, lastColumn);
        LastRow = Math.Max(firstRow, lastRow);
        LastColumn = Math.Max(firstColumn, lastColumn);
    }

    public Worksheet Sheet { get; }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int LastRow { get; }

    public int LastColumn { get; }

    /// <summary>Next area of a union, or <see langword="null"/>.</summary>
    public ReferenceValue? Next { get; private set; }

    public bool IsSingleCell => Next is null && FirstRow == LastRow && FirstColumn == LastColumn;

    public int RowCount => LastRow - FirstRow + 1;

    public int ColumnCount => LastColumn - FirstColumn + 1;

    /// <summary>
    /// Chains a freshly built area behind another and returns it, for callers assembling a
    /// multi-area value front to back — the 3-D span resolver.
    /// </summary>
    public static ReferenceValue Append(ReferenceValue tail, ReferenceValue area)
    {
        tail.Next = area;
        return area;
    }

    /// <summary>Builds a union of this reference and another.</summary>
    public static ReferenceValue Union(ReferenceValue left, ReferenceValue right)
    {
        ReferenceValue head = Copy(left);
        ReferenceValue tail = head;
        while (tail.Next is not null)
            tail = tail.Next;
        tail.Next = Copy(right);
        return head;
    }

    /// <summary>
    /// Intersects two single-area references. Returns <see langword="null"/> when the areas do
    /// not overlap, which surfaces as <c>#NULL!</c>.
    /// </summary>
    public static ReferenceValue? Intersect(ReferenceValue left, ReferenceValue right)
    {
        if (!ReferenceEquals(left.Sheet, right.Sheet))
            return null;

        int firstRow = Math.Max(left.FirstRow, right.FirstRow);
        int firstColumn = Math.Max(left.FirstColumn, right.FirstColumn);
        int lastRow = Math.Min(left.LastRow, right.LastRow);
        int lastColumn = Math.Min(left.LastColumn, right.LastColumn);
        if (firstRow > lastRow || firstColumn > lastColumn)
            return null;

        return new ReferenceValue(left.Sheet, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>The smallest range covering both corners of every area in the union.</summary>
    public ReferenceValue Bounds()
    {
        if (Next is null)
            return this;

        int firstRow = FirstRow, firstColumn = FirstColumn, lastRow = LastRow, lastColumn = LastColumn;
        for (ReferenceValue? area = Next; area is not null; area = area.Next)
        {
            firstRow = Math.Min(firstRow, area.FirstRow);
            firstColumn = Math.Min(firstColumn, area.FirstColumn);
            lastRow = Math.Max(lastRow, area.LastRow);
            lastColumn = Math.Max(lastColumn, area.LastColumn);
        }

        return new ReferenceValue(Sheet, firstRow, firstColumn, lastRow, lastColumn);
    }

    private static ReferenceValue Copy(ReferenceValue source)
    {
        var head = new ReferenceValue(source.Sheet, source.FirstRow, source.FirstColumn, source.LastRow, source.LastColumn);
        ReferenceValue tail = head;
        for (ReferenceValue? area = source.Next; area is not null; area = area.Next)
        {
            tail.Next = new ReferenceValue(area.Sheet, area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn);
            tail = tail.Next;
        }

        return head;
    }

    /// <inheritdoc/>
    public override string ToString()
    {
        var range = new RangeRef(new CellRef(FirstRow, FirstColumn), new CellRef(LastRow, LastColumn));
        return Next is null ? $"{Sheet.Name}!{range}" : $"{Sheet.Name}!{range},…";
    }
}
