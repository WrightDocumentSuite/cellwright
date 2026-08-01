namespace Cellwright;

/// <summary>
/// Describes rows or columns appearing or disappearing on one sheet, and applies that motion
/// to a reference. Insert and delete move anchored references as well as relative ones — the
/// <c>$</c> only pins a reference against copying, not against the grid moving underneath it.
/// </summary>
/// <param name="SheetName">Sheet the rows or columns were inserted into or deleted from.</param>
/// <param name="Rows">Whether rows moved; otherwise columns did.</param>
/// <param name="Index">First 1-based row or column the edit touched.</param>
/// <param name="Count">How many appeared, or a negative count of how many were removed.</param>
internal readonly record struct ReferenceShift(string SheetName, bool Rows, int Index, int Count)
{
    /// <summary>Last 1-based row or column a delete removed.</summary>
    private int RemovedEnd => Index - Count - 1;

    /// <summary>Highest coordinate the axis allows, used to clamp what an insert pushes out.</summary>
    private int Limit => Rows ? SheetLimits.MaxRow : SheetLimits.MaxColumn;

    /// <summary>Moves one range within a single sheet, or reports it as gone.</summary>
    /// <param name="range">Range to move, replaced in place.</param>
    /// <returns><see langword="false"/> when every cell of the range was deleted.</returns>
    public bool TryApply(ref RangeRef range)
    {
        int first = Rows ? range.Start.Row : range.Start.Column;
        int last = Rows ? range.End.Row : range.End.Column;
        if (!TryMove(first, last, out int movedFirst, out int movedLast))
            return false;

        range = Rows
            ? new RangeRef(new CellRef(movedFirst, range.Start.Column), new CellRef(movedLast, range.End.Column))
            : new RangeRef(new CellRef(range.Start.Row, movedFirst), new CellRef(range.End.Row, movedLast));
        return true;
    }

    /// <summary>Moves one cell, or reports it as gone.</summary>
    public bool TryApply(ref CellRef cell)
    {
        int coordinate = Rows ? cell.Row : cell.Column;
        if (!TryMove(coordinate, coordinate, out int moved, out _))
            return false;

        cell = Rows ? cell with { Row = moved } : cell with { Column = moved };
        return true;
    }

    /// <summary>
    /// Moves the two sides of a parsed A1 reference. Parts that have no coordinate on the
    /// moving axis — a column reference while rows move — are left alone.
    /// </summary>
    /// <returns>Whether the reference text has to be rewritten.</returns>
    internal bool Apply(ref RefPart start, ref RefPart end, out bool broken)
    {
        broken = false;
        if (!HasCoordinate(start.Kind))
            return false;

        int first = Coordinate(start);
        bool isRange = end.Kind != RefPartKind.None;
        int last = isRange ? Coordinate(end) : first;

        if (!TryMove(first, last, out int movedFirst, out int movedLast))
        {
            broken = true;
            return true;
        }

        if (movedFirst == first && movedLast == last)
            return false;

        start = WithCoordinate(start, movedFirst);
        if (isRange)
            end = WithCoordinate(end, movedLast);
        return true;
    }

    /// <summary>
    /// The whole rule in one place: an insert pushes down everything at or after its index and
    /// stretches a range it lands inside; a delete pulls back what survives it and clips what
    /// only partly overlaps.
    /// </summary>
    private bool TryMove(int first, int last, out int movedFirst, out int movedLast)
    {
        if (Count > 0)
        {
            movedFirst = first >= Index ? Math.Min(first + Count, Limit) : first;
            movedLast = last >= Index ? Math.Min(last + Count, Limit) : last;
            return true;
        }

        int removedEnd = RemovedEnd;
        if (first >= Index && last <= removedEnd)
        {
            movedFirst = movedLast = 0;
            return false;
        }

        movedFirst = first < Index ? first : first > removedEnd ? first + Count : Index;
        movedLast = last < Index ? last : last > removedEnd ? last + Count : Index - 1;
        return true;
    }

    private bool HasCoordinate(RefPartKind kind) => kind switch
    {
        RefPartKind.Cell => true,
        RefPartKind.Row => Rows,
        RefPartKind.Column => !Rows,
        _ => false,
    };

    private int Coordinate(RefPart part) => Rows ? part.Row : part.Column;

    private RefPart WithCoordinate(RefPart part, int value) =>
        Rows ? part with { Row = value } : part with { Column = value };
}
