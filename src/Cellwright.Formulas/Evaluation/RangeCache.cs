namespace Cellwright.Formulas;

/// <summary>Identity of a materialized rectangle: the sheet plus its corners.</summary>
internal readonly record struct RangeKey(int SheetId, int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
}

/// <summary>
/// A rectangle that has been read out of a worksheet, together with the lookup index built
/// over it on demand. The array is shared between every formula reading the same range, so
/// it must never be modified in place.
/// </summary>
internal sealed class CachedRange(RangeKey key, ArrayValue array)
{
    private readonly LookupIndex?[] _indexes = new LookupIndex?[3];
    private readonly bool[] _attempted = new bool[3];

    public RangeKey Key { get; } = key;

    public ArrayValue Array { get; } = array;

    /// <summary>Position in the cache-wide LRU chain, oldest first.</summary>
    public LinkedListNode<CachedRange>? LruNode { get; set; }

    /// <summary>Position in the owning sheet's invalidation chain.</summary>
    public LinkedListNode<CachedRange>? SheetNode { get; set; }

    /// <summary>
    /// Whether the cache holds on to this range. A transient copy is not worth indexing: the
    /// index would be built and thrown away within the same call.
    /// </summary>
    public bool IsRetained { get; set; }

    /// <summary>
    /// Hash index over one axis, built on first use. A null result is remembered too, so a
    /// range that cannot be indexed is not re-examined on every lookup.
    /// </summary>
    public LookupIndex? GetIndex(LookupAxis axis, ExcelDateSystem dateSystem)
    {
        int slot = (int)axis;
        if (!_attempted[slot])
        {
            _indexes[slot] = LookupIndex.TryBuild(Array, axis, dateSystem);
            _attempted[slot] = true;
        }

        return _indexes[slot];
    }
}

/// <summary>
/// Keeps materialized ranges alive between formulas. A worksheet full of
/// <c>VLOOKUP(..., Rates!$A$1:$B$100, ...)</c> would otherwise copy the same hundred rows once
/// per row of the report.
/// </summary>
/// <remarks>
/// Entries are dropped the moment anything is written inside them, so a stale array cannot be
/// observed. Writes are checked against a per-sheet bounding box first, which makes the common
/// case (results written far away from the lookup table) a single comparison.
/// </remarks>
internal sealed class RangeCache
{
    /// <summary>Below this size the dictionary lookup costs more than copying the cells.</summary>
    public const int MinCells = 32;

    /// <summary>
    /// How many cells the cache may hold in total. One range alone may legally reach
    /// <see cref="ArrayValue.MaxCells"/>, so without a budget the cache would trade a copy
    /// per call for gigabytes of retained memory.
    /// </summary>
    public const long MaxCells = 2_000_000;

    private readonly long _maxCells;
    private readonly Dictionary<RangeKey, CachedRange> _entries = [];
    private readonly Dictionary<int, SheetEntries> _bySheet = [];
    private readonly LinkedList<CachedRange> _lru = [];

    public RangeCache(long maxCells = MaxCells)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(maxCells);
        _maxCells = maxCells;
    }

    /// <summary>Number of reads served from the cache.</summary>
    public int Hits { get; private set; }

    /// <summary>Number of reads that had to materialize the range.</summary>
    public int Misses { get; private set; }

    /// <summary>Cells currently retained.</summary>
    public long RetainedCells { get; private set; }

    /// <summary>Number of retained ranges.</summary>
    public int Count => _entries.Count;

    /// <summary>Looks up a range, counting the hit or miss.</summary>
    public bool TryGet(in RangeKey key, out CachedRange entry)
    {
        if (_entries.TryGetValue(key, out CachedRange? found))
        {
            Touch(found);
            Hits++;
            entry = found;
            return true;
        }

        Misses++;
        entry = null!;
        return false;
    }

    /// <summary>Retains a materialized range, evicting older ones if the budget is exceeded.</summary>
    public void Add(CachedRange entry)
    {
        long cells = entry.Array.Count;
        if (cells > _maxCells)
            return;

        if (_entries.TryGetValue(entry.Key, out CachedRange? previous))
            Remove(previous);

        entry.IsRetained = true;
        _entries.Add(entry.Key, entry);
        RetainedCells += cells;

        entry.LruNode = _lru.AddLast(entry);

        SheetEntries sheet = GetSheet(entry.Key.SheetId);
        entry.SheetNode = sheet.Ranges.AddLast(entry);
        sheet.Include(entry.Key);

        while (RetainedCells > _maxCells)
        {
            if (!TryEvictLeastRecentlyUsed())
                break;
        }
    }

    /// <summary>
    /// Drops every range covering the written cell. Called for both the engine's own results
    /// and edits made through the model, which is what keeps cached arrays honest.
    /// </summary>
    public void Invalidate(Worksheet sheet, int row, int column)
    {
        if (_entries.Count == 0 || !_bySheet.TryGetValue(sheet.Id, out SheetEntries? entries))
            return;
        if (!entries.MayContain(row, column))
            return;

        LinkedListNode<CachedRange>? node = entries.Ranges.First;
        while (node is not null)
        {
            LinkedListNode<CachedRange>? next = node.Next;
            CachedRange entry = node.Value;
            if (entry.Key.Contains(row, column))
                Remove(entry);

            node = next;
        }
    }

    /// <summary>Drops everything; used when the workbook structure changes under the engine.</summary>
    public void Clear()
    {
        foreach (CachedRange entry in _entries.Values)
        {
            entry.LruNode = null;
            entry.SheetNode = null;
            entry.IsRetained = false;
        }

        _entries.Clear();
        _bySheet.Clear();
        _lru.Clear();
        RetainedCells = 0;
    }

    /// <summary>Resets the hit and miss counters between measurements.</summary>
    public void ResetCounters()
    {
        Hits = 0;
        Misses = 0;
    }

    private bool TryEvictLeastRecentlyUsed()
    {
        LinkedListNode<CachedRange>? oldest = _lru.First;
        if (oldest is null)
            return false;

        Remove(oldest.Value);
        return true;
    }

    private void Remove(CachedRange entry)
    {
        if (!_entries.TryGetValue(entry.Key, out CachedRange? retained) ||
            !ReferenceEquals(retained, entry))
        {
            return;
        }

        _entries.Remove(entry.Key);
        RetainedCells -= entry.Array.Count;

        if (entry.LruNode is not null)
            _lru.Remove(entry.LruNode);

        if (entry.SheetNode is not null &&
            _bySheet.TryGetValue(entry.Key.SheetId, out SheetEntries? sheet))
        {
            sheet.Ranges.Remove(entry.SheetNode);
            if (sheet.Ranges.Count == 0)
                _bySheet.Remove(entry.Key.SheetId);
        }

        entry.LruNode = null;
        entry.SheetNode = null;
        entry.IsRetained = false;
    }

    private void Touch(CachedRange entry)
    {
        LinkedListNode<CachedRange> node = entry.LruNode!;
        if (!ReferenceEquals(node, _lru.Last))
        {
            _lru.Remove(node);
            _lru.AddLast(node);
        }
    }

    private SheetEntries GetSheet(int sheetId)
    {
        if (!_bySheet.TryGetValue(sheetId, out SheetEntries? entries))
            _bySheet[sheetId] = entries = new SheetEntries();
        return entries;
    }

    /// <summary>Ranges retained for one sheet, plus the box that encloses all of them.</summary>
    private sealed class SheetEntries
    {
        private int _firstRow = int.MaxValue;
        private int _firstColumn = int.MaxValue;
        private int _lastRow = int.MinValue;
        private int _lastColumn = int.MinValue;

        public LinkedList<CachedRange> Ranges { get; } = [];

        public bool MayContain(int row, int column) =>
            row >= _firstRow && row <= _lastRow && column >= _firstColumn && column <= _lastColumn;

        public void Include(in RangeKey key)
        {
            _firstRow = Math.Min(_firstRow, key.FirstRow);
            _firstColumn = Math.Min(_firstColumn, key.FirstColumn);
            _lastRow = Math.Max(_lastRow, key.LastRow);
            _lastColumn = Math.Max(_lastColumn, key.LastColumn);
        }

    }
}
