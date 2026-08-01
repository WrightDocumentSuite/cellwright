using System.Collections;
using Cellwright.Model;
using Cellwright.Streaming;

namespace Cellwright;

/// <summary>
/// The ordered worksheets of a <see cref="Workbook"/>.
/// </summary>
public sealed class SheetCollection : IReadOnlyList<Worksheet>
{
    private readonly Workbook _workbook;
    private readonly List<Worksheet> _sheets = [];

    internal SheetCollection(Workbook workbook) => _workbook = workbook;

    /// <inheritdoc/>
    public int Count => _sheets.Count;

    /// <inheritdoc/>
    public Worksheet this[int index] => _sheets[index];

    /// <summary>Finds a sheet by name (case-insensitive, like Excel).</summary>
    /// <exception cref="KeyNotFoundException">No sheet has that name.</exception>
    public Worksheet this[string name] =>
        Find(name) ?? throw new KeyNotFoundException($"The workbook has no sheet named '{name}'.");

    /// <summary>Adds a new empty sheet at the end of the workbook.</summary>
    public Worksheet Add(string name)
    {
        SheetNameValidator.Validate(name);
        if (Find(name) is not null)
            throw new ArgumentException($"A sheet named '{name}' already exists.", nameof(name));

        var sheet = new Worksheet(_workbook, name, _workbook.NextSheetId++);
        _sheets.Add(sheet);
        _workbook.IsFullyCalculated = false;
        return sheet;
    }

    /// <summary>Removes a sheet by name. Returns whether a sheet was removed.</summary>
    public bool Remove(string name)
    {
        if (Find(name) is not { } sheet || !_sheets.Remove(sheet))
            return false;

        // Formulas reference sheets by identity, so removal turns them into #REF! on the next
        // calculation rather than silently binding to whatever sheet took the same name.
        _workbook.IsFullyCalculated = false;
        return true;
    }

    /// <summary>
    /// Moves a sheet to another position in the workbook, sliding the sheets in between along.
    /// </summary>
    /// <param name="sheet">Sheet to move; it must belong to this workbook.</param>
    /// <param name="index">Zero-based position it should end up at.</param>
    public void Move(Worksheet sheet, int index)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        int from = _sheets.IndexOf(sheet);
        if (from < 0)
            throw new ArgumentException($"Sheet '{sheet.Name}' does not belong to this workbook.", nameof(sheet));

        ArgumentOutOfRangeException.ThrowIfNegative(index);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets.Count);
        if (from == index)
            return;

        _sheets.RemoveAt(from);
        _sheets.Insert(index, sheet);
    }

    /// <summary>Moves a sheet found by name.</summary>
    /// <param name="name">Name of the sheet to move.</param>
    /// <param name="index">Zero-based position it should end up at.</param>
    public void Move(string name, int index) =>
        Move(this[name], index);

    /// <summary>
    /// Duplicates a sheet under a new name, placing the copy right after the original.
    /// </summary>
    /// <param name="sheet">Sheet to duplicate; it must belong to this workbook.</param>
    /// <param name="newName">Name for the copy; it must not be in use.</param>
    /// <remarks>
    /// Cells, styles, formulas, layout, notes, links, rules, print settings and protection are
    /// all copied. What the copy does not get is the part of the source that only lived in the
    /// file it was loaded from — drawings, pictures and anything else the model carries as
    /// untouched markup, which belongs to one sheet part and cannot be in two.
    /// </remarks>
    public Worksheet Copy(Worksheet sheet, string newName)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        int from = _sheets.IndexOf(sheet);
        if (from < 0)
            throw new ArgumentException($"Sheet '{sheet.Name}' does not belong to this workbook.", nameof(sheet));

        Worksheet copy = Add(newName);
        WorksheetCopier.Fill(sheet, copy);
        Move(copy, from + 1);
        return copy;
    }

    /// <summary>Duplicates a sheet found by name.</summary>
    /// <param name="name">Name of the sheet to duplicate.</param>
    /// <param name="newName">Name for the copy; it must not be in use.</param>
    public Worksheet Copy(string name, string newName) => Copy(this[name], newName);

    /// <summary>Finds a sheet by its stable identifier, or <see langword="null"/> when removed.</summary>
    internal Worksheet? FindById(int id)
    {
        foreach (Worksheet sheet in _sheets)
        {
            if (sheet.Id == id)
                return sheet;
        }

        return null;
    }

    /// <summary>Finds a sheet by name, or <see langword="null"/>.</summary>
    public Worksheet? Find(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        foreach (Worksheet sheet in _sheets)
        {
            if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                return sheet;
        }

        return null;
    }

    internal bool NameInUseByOther(Worksheet owner, string name) =>
        Find(name) is { } existing && !ReferenceEquals(existing, owner);

    /// <inheritdoc/>
    public IEnumerator<Worksheet> GetEnumerator() => _sheets.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
