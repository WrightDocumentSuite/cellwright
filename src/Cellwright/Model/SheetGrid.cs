namespace Cellwright.Model;

/// <summary>
/// Data-oriented sparse grid: one paged column store per aspect (values, style indexes,
/// formulas), no per-cell objects. A cell is "present" when its slot differs from
/// <c>default(T)</c>; pages of 512 rows are allocated lazily per column.
/// </summary>
internal sealed class SheetGrid<T>
{
    private const int PageBits = 9;
    private const int PageSize = 1 << PageBits;   // 512 rows per page.
    private const int PageMask = PageSize - 1;

    private ColumnStore?[] _columns = [];

    /// <summary>Highest 1-based column that ever received a value (upper bound).</summary>
    public int MaxColumn { get; private set; }

    /// <summary>Highest 1-based row that ever received a value (upper bound).</summary>
    public int MaxRow { get; private set; }

    /// <summary>Reads a slot; absent cells return <c>default</c>.</summary>
    public T Get(int row, int column)
    {
        if (column > _columns.Length || _columns[column - 1] is not { } store)
            return default!;
        return store.Get(row);
    }

    /// <summary>Writes a slot. Writing <c>default</c> clears it.</summary>
    public void Set(int row, int column, T value)
    {
        if (!SheetLimits.IsValidRow(row) || !SheetLimits.IsValidColumn(column))
            throw new ArgumentOutOfRangeException(nameof(row), $"Cell R{row}C{column} lies outside the worksheet grid.");

        if (column > _columns.Length)
        {
            if (EqualityComparer<T>.Default.Equals(value, default!))
                return;
            Array.Resize(ref _columns, Math.Max(Math.Max(_columns.Length * 2, column), 8));
        }

        (_columns[column - 1] ??= new ColumnStore()).Set(row, value);
        if (!EqualityComparer<T>.Default.Equals(value, default!))
        {
            MaxColumn = Math.Max(MaxColumn, column);
            MaxRow = Math.Max(MaxRow, row);
        }
    }

    /// <summary>Fills a buffer with the slots of one row from column 1 to <paramref name="columnCount"/>.</summary>
    public void ReadRow(int row, Span<T> destination, int columnCount)
    {
        int limit = Math.Min(columnCount, _columns.Length);
        for (int column = 1; column <= limit; column++)
            destination[column - 1] = _columns[column - 1] is { } store ? store.Get(row) : default!;
        if (limit < columnCount)
            destination[limit..columnCount].Clear();
    }

    /// <summary>
    /// Walks every occupied slot in column-major order. Empty pages are skipped, so the cost
    /// tracks the number of values rather than the size of the grid.
    /// </summary>
    public IEnumerable<(int Row, int Column, T Value)> EnumerateOccupied()
    {
        var comparer = EqualityComparer<T>.Default;
        for (int column = 1; column <= _columns.Length; column++)
        {
            if (_columns[column - 1] is not { } store)
                continue;

            foreach ((int row, T value) in store.Enumerate())
            {
                if (!comparer.Equals(value, default!))
                    yield return (row, column, value);
            }
        }
    }

    /// <summary>
    /// Moves rows to make room for <paramref name="count"/> new ones at <paramref name="index"/>,
    /// or closes the gap left by deleting that many. Cells pushed past the last row of the grid
    /// are dropped, the way Excel drops them once it has warned about it.
    /// </summary>
    public void ShiftRows(int index, int count)
    {
        if (count == 0)
            return;

        int maxRow = MaxRow;
        foreach (ColumnStore? store in _columns)
            store?.ShiftRows(index, count, maxRow);

        MaxRow = count > 0
            ? Math.Min(maxRow + count, SheetLimits.MaxRow)
            : Math.Max(0, Math.Min(maxRow, maxRow + count));
    }

    /// <summary>Column-wise counterpart of <see cref="ShiftRows"/>; whole slices move at once.</summary>
    public void ShiftColumns(int index, int count)
    {
        if (count == 0)
            return;

        if (count > 0)
        {
            int last = Math.Min(MaxColumn + count, SheetLimits.MaxColumn);
            EnsureColumnCapacity(last);
            for (int column = last; column >= index + count; column--)
                _columns[column - 1] = _columns[column - count - 1];
            for (int column = index; column < index + count && column <= last; column++)
                _columns[column - 1] = null;
            MaxColumn = last;
            return;
        }

        int removed = -count;
        for (int column = index; column <= MaxColumn; column++)
        {
            int source = column + removed;
            _columns[column - 1] = source <= _columns.Length ? _columns[source - 1] : null;
        }

        MaxColumn = Math.Max(0, MaxColumn - removed);
    }

    private void EnsureColumnCapacity(int columns)
    {
        if (columns > _columns.Length)
            Array.Resize(ref _columns, columns);
    }

    /// <summary>
    /// Rewrites every occupied slot in place. Nothing is allocated or moved, so this stays
    /// cheap on a sheet with a million formulas to retarget.
    /// </summary>
    public void Transform(Func<T, T> transform)
    {
        var comparer = EqualityComparer<T>.Default;
        foreach (ColumnStore? store in _columns)
            store?.Transform(transform, comparer);
    }

    /// <summary>Whether any slot in the given row differs from <c>default</c>.</summary>
    public bool IsRowOccupied(int row)
    {
        foreach (ColumnStore? store in _columns)
        {
            if (store is not null && !EqualityComparer<T>.Default.Equals(store.Get(row), default!))
                return true;
        }

        return false;
    }

    private sealed class ColumnStore
    {
        private T[]?[] _pages = [];

        public T Get(int row)
        {
            int pageIndex = (row - 1) >> PageBits;
            if (pageIndex >= _pages.Length || _pages[pageIndex] is not { } page)
                return default!;
            return page[(row - 1) & PageMask];
        }

        public void Set(int row, T value)
        {
            int pageIndex = (row - 1) >> PageBits;
            if (pageIndex >= _pages.Length)
            {
                if (EqualityComparer<T>.Default.Equals(value, default!))
                    return;
                Array.Resize(ref _pages, Math.Max(Math.Max(_pages.Length * 2, pageIndex + 1), 4));
            }

            T[]? page = _pages[pageIndex];
            if (page is null)
            {
                if (EqualityComparer<T>.Default.Equals(value, default!))
                    return;
                _pages[pageIndex] = page = new T[PageSize];
            }

            page[(row - 1) & PageMask] = value;
        }

        /// <summary>
        /// Moves the rows of this column. The copy runs away from the gap — downward for an
        /// insert, upward for a delete — so a slot is only read before it is overwritten.
        /// </summary>
        public void ShiftRows(int index, int count, int maxRow)
        {
            if (count > 0)
            {
                for (int row = Math.Min(maxRow, SheetLimits.MaxRow - count); row >= index; row--)
                    Set(row + count, Get(row));
                for (int row = index; row < index + count && row <= maxRow; row++)
                    Set(row, default!);
                return;
            }

            int removed = -count;
            for (int row = index; row <= maxRow; row++)
                Set(row, Get(row + removed));
        }

        public void Transform(Func<T, T> transform, EqualityComparer<T> comparer)
        {
            foreach (T[]? page in _pages)
            {
                if (page is null)
                    continue;

                for (int slot = 0; slot < page.Length; slot++)
                {
                    if (!comparer.Equals(page[slot], default!))
                        page[slot] = transform(page[slot]);
                }
            }
        }

        public IEnumerable<(int Row, T Value)> Enumerate()
        {
            for (int pageIndex = 0; pageIndex < _pages.Length; pageIndex++)
            {
                if (_pages[pageIndex] is not { } page)
                    continue;

                for (int slot = 0; slot < page.Length; slot++)
                    yield return ((pageIndex << PageBits) + slot + 1, page[slot]);
            }
        }
    }
}
