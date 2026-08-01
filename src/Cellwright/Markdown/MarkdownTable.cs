using System.Text;
using Cellwright.Styles;

namespace Cellwright.Markdown;

/// <summary>
/// Writes the used range of one sheet as a GitHub Flavored Markdown table.
/// </summary>
/// <remarks>
/// The single hard rule of the extension is that the header row and the delimiter row hold the
/// same number of cells, or the block is not a table at all — so both are emitted from the same
/// column count and never from what the rows happen to contain.
/// </remarks>
internal static class MarkdownTable
{
    /// <summary>Writes the table, followed by the blank line that closes the block.</summary>
    public static void Write(StringBuilder builder, Worksheet sheet, RangeRef used, MarkdownExportOptions options)
    {
        var overlay = new SheetOverlay(sheet);
        int columns = used.End.Column;
        int firstDataRow = used.Start.Row;

        if (options.HeaderRow == MarkdownHeaderRow.FirstRow)
        {
            WriteRow(builder, sheet, used.Start.Row, columns, options, overlay);
            firstDataRow++;
        }
        else
        {
            WriteLetterHeader(builder, columns);
        }

        WriteDelimiter(builder, sheet, used.Start.Row, columns);
        for (int row = firstDataRow; row <= used.End.Row; row++)
            WriteRow(builder, sheet, row, columns, options, overlay);

        builder.Append('\n');
    }

    private static void WriteLetterHeader(StringBuilder builder, int columns)
    {
        builder.Append('|');
        for (int column = 1; column <= columns; column++)
            builder.Append(' ').Append(ColumnName.Format(column)).Append(" |");
        builder.Append('\n');
    }

    /// <summary>
    /// Writes the row that carries the column alignments. They are read off the first row of the
    /// used range: with a header that is the header cell, and without one it is still the cell
    /// whose look the column takes.
    /// </summary>
    private static void WriteDelimiter(StringBuilder builder, Worksheet sheet, int alignmentRow, int columns)
    {
        builder.Append('|');
        for (int column = 1; column <= columns; column++)
        {
            HorizontalAlignment alignment = sheet.Cell(alignmentRow, column).StyleValue.Alignment.Horizontal;
            builder.Append(' ').Append(Delimiter(alignment)).Append(" |");
        }

        builder.Append('\n');
    }

    private static string Delimiter(HorizontalAlignment alignment) => alignment switch
    {
        HorizontalAlignment.Left => ":---",
        HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous => ":---:",
        HorizontalAlignment.Right => "---:",
        _ => "---",
    };

    private static void WriteRow(
        StringBuilder builder, Worksheet sheet, int row, int columns, MarkdownExportOptions options, SheetOverlay overlay)
    {
        overlay.MoveTo(row);
        builder.Append('|');
        for (int column = 1; column <= columns; column++)
        {
            builder.Append(' ');
            WriteCell(builder, sheet, new CellRef(row, column), options, overlay);
            builder.Append(" |");
        }

        builder.Append('\n');
    }

    /// <summary>
    /// Writes one cell as Excel shows it: the number format applied to the value, the cached
    /// result of a formula rather than its text, and nothing at all for the cells a merge
    /// swallows — Markdown has no way to span columns.
    /// </summary>
    private static void WriteCell(
        StringBuilder builder, Worksheet sheet, CellRef reference, MarkdownExportOptions options, SheetOverlay overlay)
    {
        if (overlay.IsMergedAway(reference))
            return;

        Cell cell = sheet.Cell(reference);
        string rendered = NumberFormatRenderer.Format(
            cell.Value, cell.StyleValue.NumberFormat, options.Culture, sheet.Workbook.DateSystem);

        if (overlay.LinkAt(reference) is not { } link)
        {
            MarkdownText.AppendCell(builder, rendered);
            return;
        }

        string destination = LinkDestination(link);
        string text = LinkText(rendered, link, destination);
        if (MarkdownText.IsSafeDestination(destination))
            MarkdownText.AppendLink(builder, text, destination);
        else
            MarkdownText.AppendCell(builder, text);
    }

    /// <summary>
    /// What a linked cell reads as. The cell's own value comes first because that is what Excel
    /// shows; the <c>display</c> attribute beside the link is a cached copy Excel itself ignores,
    /// and is only worth reading when the cell under the link is blank.
    /// </summary>
    private static string LinkText(string rendered, Hyperlink link, string destination)
    {
        if (rendered.Length > 0)
            return rendered;
        if (link.Display is { Length: > 0 } display)
            return display;
        return destination;
    }

    /// <summary>
    /// An external hyperlink may also name a location inside its target. Excel stores the two
    /// parts separately; a Markdown link carries the location as the target's fragment.
    /// </summary>
    private static string LinkDestination(Hyperlink link)
    {
        if (link.Target is not { } target)
            return link.Location ?? "";
        return link.Location is { Length: > 0 } location ? $"{target}#{location}" : target;
    }

    /// <summary>
    /// The merges and hyperlinks of a sheet, narrowed to one row at a time. Both are flat lists
    /// on the sheet, and rescanning them for every cell of a large export would cost more than
    /// everything else put together.
    /// </summary>
    private sealed class SheetOverlay
    {
        private readonly IReadOnlyList<RangeRef> _merges;
        private readonly HyperlinkCollection _links;
        private readonly List<RangeRef> _rowMerges = [];
        private readonly List<Hyperlink> _rowLinks = [];

        public SheetOverlay(Worksheet sheet)
        {
            _merges = sheet.MergedRanges;
            _links = sheet.Hyperlinks;
        }

        /// <summary>Collects what covers <paramref name="row"/>; a sheet with neither pays nothing.</summary>
        public void MoveTo(int row)
        {
            if (_merges.Count > 0)
            {
                _rowMerges.Clear();
                foreach (RangeRef merge in _merges)
                {
                    if (merge.Start.Row <= row && row <= merge.End.Row)
                        _rowMerges.Add(merge);
                }
            }

            if (_links.Count == 0)
                return;

            _rowLinks.Clear();
            foreach (Hyperlink link in _links)
            {
                if (link.Range.Start.Row <= row && row <= link.Range.End.Row)
                    _rowLinks.Add(link);
            }
        }

        /// <summary>Whether the cell belongs to a merge whose value another cell already carries.</summary>
        public bool IsMergedAway(CellRef cell)
        {
            foreach (RangeRef merge in _rowMerges)
            {
                if (merge.Contains(cell) && merge.Start != cell)
                    return true;
            }

            return false;
        }

        /// <summary>The link covering the cell, or <see langword="null"/>.</summary>
        public Hyperlink? LinkAt(CellRef cell)
        {
            foreach (Hyperlink link in _rowLinks)
            {
                if (link.Range.Contains(cell))
                    return link;
            }

            return null;
        }
    }
}
