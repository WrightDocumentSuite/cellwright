using System.Collections;

namespace Cellwright;

/// <summary>
/// A named table over a rectangle of cells — what Excel calls "Format as Table". The model
/// covers the range, the column names and the banded-style flags; filter state and calculated
/// column formulas stay in the file's own markup and are carried through untouched.
/// </summary>
public sealed class ExcelTable
{
    /// <summary>Table name, unique within the workbook and usable in structured references.</summary>
    public required string Name { get; init; }

    private RangeRef _range;
    private IReadOnlyList<string> _columns = [];

    /// <summary>The collection told about edits, so a loaded table is not saved as it was read.</summary>
    internal TableCollection? Owner { get; set; }

    /// <summary>Cells the table covers, header row included.</summary>
    /// <exception cref="ArgumentException">The new range is not as wide as <see cref="Columns"/>.</exception>
    public required RangeRef Range
    {
        get => _range;
        set
        {
            if (_range == value)
                return;

            if (Owner is not null && value.ColumnCount != _columns.Count)
            {
                throw new ArgumentException(
                    $"Table '{Name}' has {_columns.Count} column(s), so its range cannot become {value.ColumnCount} wide.",
                    nameof(value));
            }

            _range = value;
            Owner?.MarkModified();
        }
    }

    /// <summary>Column names, left to right. One per column of <see cref="Range"/>.</summary>
    public required IReadOnlyList<string> Columns
    {
        get => _columns;
        init => _columns = value;
    }

    /// <summary>Changes the range and the columns together, as a structural edit has to.</summary>
    internal void Reshape(RangeRef range, IReadOnlyList<string> columns)
    {
        _range = range;
        _columns = columns;
    }

    /// <summary>Whether the first row of the range holds the column names.</summary>
    public bool HasHeaderRow { get; init; } = true;

    /// <summary>Whether the last row of the range holds totals.</summary>
    public bool HasTotalsRow { get; init; }

    /// <summary>Name of the built-in table style, e.g. <c>TableStyleMedium2</c>.</summary>
    public string? StyleName { get; init; } = "TableStyleMedium2";

    /// <summary>Whether rows are banded.</summary>
    public bool ShowRowStripes { get; init; } = true;

    /// <summary>Whether columns are banded.</summary>
    public bool ShowColumnStripes { get; init; }

    /// <summary>Display name shown in the UI; defaults to <see cref="Name"/>.</summary>
    public string? DisplayName { get; init; }

    /// <summary>Identifier within the workbook; assigned on save when the table is new.</summary>
    internal int Id { get; set; }

    /// <summary>Whether the loaded table part carried filter criteria, hiding the rows they reject.</summary>
    internal bool FilterIsActive { get; init; }

    /// <inheritdoc/>
    public override string ToString() => $"{Name} over {Range}";
}

/// <summary>The tables of one worksheet.</summary>
public sealed class TableCollection : IReadOnlyList<ExcelTable>
{
    private readonly Worksheet _sheet;
    private readonly List<ExcelTable> _tables = [];

    internal TableCollection(Worksheet sheet) => _sheet = sheet;

    /// <inheritdoc/>
    public int Count => _tables.Count;

    /// <inheritdoc/>
    public ExcelTable this[int index] => _tables[index];

    /// <summary>
    /// Whether the model, rather than the loaded file, owns the sheet's tables. Until
    /// something changes, the original table parts are copied through with their filter state
    /// and calculated columns intact.
    /// </summary>
    internal bool IsModified { get; private set; }

    /// <summary>Finds a table by name (case-insensitive), or <see langword="null"/>.</summary>
    public ExcelTable? Find(string name)
    {
        foreach (ExcelTable table in _tables)
        {
            if (string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase))
                return table;
        }

        return null;
    }

    /// <summary>
    /// Creates a table over a range, taking the column names from its header row. Cells of
    /// the header row that are blank get a generated name, as Excel does.
    /// </summary>
    public ExcelTable Add(string name, RangeRef range, bool hasHeaderRow = true)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (Find(name) is not null)
            throw new ArgumentException($"A table named '{name}' already exists on sheet '{_sheet.Name}'.", nameof(name));

        return Add(new ExcelTable
        {
            Name = name,
            Range = range,
            Columns = ReadColumnNames(range, hasHeaderRow),
            HasHeaderRow = hasHeaderRow,
        });
    }

    /// <summary>Creates a table over an A1 range like <c>A1:D100</c>.</summary>
    public ExcelTable Add(string name, string range, bool hasHeaderRow = true) =>
        Add(name, RangeRef.Parse(range), hasHeaderRow);

    /// <summary>Adds a fully described table.</summary>
    public ExcelTable Add(ExcelTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        if (table.Columns.Count != table.Range.ColumnCount)
        {
            throw new ArgumentException(
                $"Table '{table.Name}' declares {table.Columns.Count} column(s) over a range {table.Range.ColumnCount} wide.",
                nameof(table));
        }

        table.Owner = this;
        _tables.Add(table);
        IsModified = true;
        return table;
    }

    /// <summary>Removes a table by name. Returns whether one was removed.</summary>
    public bool Remove(string name)
    {
        if (Find(name) is not { } table)
            return false;

        _tables.Remove(table);
        table.Owner = null;
        IsModified = true;
        return true;
    }

    /// <summary>Adds a table read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(ExcelTable table)
    {
        table.Owner = this;
        _tables.Add(table);
    }

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit moved a table.</summary>
    internal void MarkModified() => IsModified = true;

    /// <inheritdoc/>
    public IEnumerator<ExcelTable> GetEnumerator() => _tables.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private List<string> ReadColumnNames(RangeRef range, bool hasHeaderRow)
    {
        List<string> names = new(range.ColumnCount);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < range.ColumnCount; i++)
        {
            string name = hasHeaderRow
                ? _sheet[range.Start.Row, range.Start.Column + i].ToString()
                : "";
            if (string.IsNullOrWhiteSpace(name))
                name = $"Column{i + 1}";
            for (int suffix = 2; !used.Add(name); suffix++)
                name = $"Column{i + 1}_{suffix}";
            names.Add(name);
        }

        return names;
    }
}
