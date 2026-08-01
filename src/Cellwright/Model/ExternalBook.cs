namespace Cellwright;

/// <summary>
/// Another workbook as this one remembers it: the path it was linked from, its sheet names,
/// and the cell values Excel cached the last time the link was refreshed (§18.14). Formulas
/// that read <c>[1]Sheet1!A1</c> answer from here — the other file itself is never opened,
/// exactly as Excel behaves while the source is closed.
/// </summary>
public sealed class ExternalBook
{
    internal ExternalBook(string? targetPath, IReadOnlyList<string> sheetNames, IReadOnlyList<ExternalSheetCache> sheets)
    {
        TargetPath = targetPath;
        SheetNames = sheetNames;
        Sheets = sheets;
    }

    /// <summary>Where the link points, as the relationship recorded it.</summary>
    public string? TargetPath { get; }

    /// <summary>Sheet names of the external book, in its own order.</summary>
    public IReadOnlyList<string> SheetNames { get; }

    internal IReadOnlyList<ExternalSheetCache> Sheets { get; }

    /// <summary>Index of a sheet by name, the way external references name them, or -1.</summary>
    internal int IndexOfSheet(string name)
    {
        for (int i = 0; i < SheetNames.Count; i++)
        {
            if (string.Equals(SheetNames[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }

    /// <summary>The cached value of one cell; blank when the cache never held it.</summary>
    public CellValue GetCachedValue(string sheetName, int row, int column)
    {
        int index = IndexOfSheet(sheetName);
        return index < 0 ? CellValue.Blank : Sheets[index].Read(row, column);
    }
}

/// <summary>The cached cells of one external sheet, sparse, with the extent they reach.</summary>
internal sealed class ExternalSheetCache
{
    private readonly Dictionary<long, CellValue> _cells = [];

    /// <summary>Highest cached row, for clamping whole-column references.</summary>
    public int LastRow { get; private set; }

    /// <summary>Highest cached column, for clamping whole-row references.</summary>
    public int LastColumn { get; private set; }

    public void Add(int row, int column, in CellValue value)
    {
        _cells[Key(row, column)] = value;
        LastRow = Math.Max(LastRow, row);
        LastColumn = Math.Max(LastColumn, column);
    }

    public CellValue Read(int row, int column) =>
        _cells.TryGetValue(Key(row, column), out CellValue value) ? value : CellValue.Blank;

    private static long Key(int row, int column) => ((long)row << 20) | (uint)column;
}
