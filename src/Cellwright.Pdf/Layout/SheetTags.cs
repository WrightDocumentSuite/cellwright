namespace Cellwright.Pdf.Layout;

/// <summary>
/// The structure elements of one sheet, when the export is tagged: a sheet is a table, each of
/// its printed rows a row of that table, and each cell a cell of that row.
/// </summary>
/// <remarks>
/// Elements are named by what they belong to rather than by where they are drawn, so a title row
/// repeated on ten pages asks for the same element ten times and stays one row of one table.
/// A row of the repeated titles is a header row, which is what lets a reader say which column a
/// figure came from.
/// </remarks>
internal sealed class SheetTags
{
    private readonly HashSet<int> _titleRows;
    private readonly TagRef _table;
    private readonly Dictionary<int, TagRef> _rows = [];
    private readonly Dictionary<(int Row, int Column), TagRef> _cells = [];

    internal SheetTags(Worksheet sheet, IEnumerable<int> titleRows)
    {
        _titleRows = [.. titleRows];
        _table = new TagRef("Table", sheet);
    }

    /// <summary>
    /// The element a cell belongs to. The same cell asked for twice — a title row is drawn on
    /// every page — answers the same element, which is what keeps the tree one table.
    /// </summary>
    /// <param name="row">The row number, 1-based.</param>
    /// <param name="column">The column number, 1-based.</param>
    public TagRef Cell(int row, int column)
    {
        if (_cells.TryGetValue((row, column), out TagRef? cached))
            return cached;

        if (!_rows.TryGetValue(row, out TagRef? rowTag))
        {
            rowTag = new TagRef("TR", new object()) { Parent = _table };
            _rows[row] = rowTag;
        }

        var cell = new TagRef(_titleRows.Contains(row) ? "TH" : "TD", new object()) { Parent = rowTag };
        _cells[(row, column)] = cell;
        return cell;
    }
}
