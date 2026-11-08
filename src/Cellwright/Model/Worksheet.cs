using Cellwright.IO;
using Cellwright.Model;
using Cellwright.Streaming;

namespace Cellwright;

/// <summary>
/// One worksheet of an in-memory <see cref="Workbook"/>. Cell data lives in sparse
/// column slices; <see cref="Cellwright.Cell"/> and <see cref="CellRange"/> are
/// lightweight struct facades over this storage.
/// </summary>
public sealed class Worksheet
{
    private string _name;
    private RangeRef? _autoFilterRange;
    private List<RangeRef>? _filteredAreas;

    internal Worksheet(Workbook workbook, string name, int id)
    {
        Workbook = workbook;
        _name = name;
        Id = id;
        Charts = new ChartCollection(this);
        ThreadedComments = new ThreadedCommentCollection(this);
        Tables = new TableCollection(this);
    }

    /// <summary>The owning workbook.</summary>
    public Workbook Workbook { get; }

    /// <summary>
    /// Identity of the sheet within its workbook, stable across renames. Formula programs
    /// bind to this rather than to the name, so renaming a sheet keeps formulas valid.
    /// </summary>
    internal int Id { get; }

    /// <summary>
    /// Sheet name; renaming validates Excel's rules and uniqueness, then retargets everything
    /// that named the sheet: formulas across the workbook, defined names, input rules,
    /// conditional formats, internal hyperlinks and the charts of a preserved package.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            SheetNameValidator.Validate(value);
            if (Workbook.Sheets.NameInUseByOther(this, value))
                throw new ArgumentException($"A sheet named '{value}' already exists.", nameof(value));
            if (string.Equals(_name, value, StringComparison.Ordinal))
                return;

            string previous = _name;
            _name = value;
            Workbook.RetargetSheetName(previous, value);
        }
    }

    /// <summary>How visible the sheet is in the Excel UI.</summary>
    public SheetVisibility Visibility { get; set; }

    /// <summary>
    /// Whether the sheet is hidden at all. Setting it picks the ordinary hidden state; use
    /// <see cref="Visibility"/> to tell that state apart from very hidden.
    /// </summary>
    public bool IsHidden
    {
        get => Visibility != SheetVisibility.Visible;
        set => Visibility = value ? SheetVisibility.Hidden : SheetVisibility.Visible;
    }

    /// <summary>Number of top rows frozen when scrolling.</summary>
    public int FrozenRows { get; set; }

    /// <summary>Number of left columns frozen when scrolling.</summary>
    public int FrozenColumns { get; set; }

    /// <summary>Color of the sheet tab; <see cref="ExcelColor.Auto"/> for the default.</summary>
    public ExcelColor TabColor { get; set; } = ExcelColor.Auto;

    /// <summary>Whether the grid lines are drawn.</summary>
    public bool ShowGridLines { get; set; } = true;

    /// <summary>Zoom percentage of the sheet view; 100 unless the file said otherwise.</summary>
    public int ZoomScale { get; set; } = 100;

    /// <summary>Hyperlinks placed over cells of this sheet.</summary>
    public HyperlinkCollection Hyperlinks { get; } = new();

    /// <summary>Input rules attached to ranges of this sheet.</summary>
    public DataValidationCollection Validations { get; } = new();

    /// <summary>Conditional formatting rules of this sheet, in the order they are tried.</summary>
    public ConditionalFormatCollection ConditionalFormats { get; } = new();

    /// <summary>How the sheet is printed: paper, orientation, scaling, margins, headers, breaks.</summary>
    public PageSetup PageSetup { get; } = new();

    /// <summary>What stays possible once the sheet is protected. Off until <c>Protect</c> is called.</summary>
    public SheetProtection Protection { get; } = new();

    /// <summary>Notes attached to cells of this sheet.</summary>
    public CommentCollection Comments { get; } = new();

    /// <summary>
    /// The charts embedded on this sheet: the ones the file carried and the ones added through
    /// <see cref="AddChart(ChartKind, string, string, string?, string?, string?)"/>. A loaded
    /// chart's raw markup is the source of truth and writes back byte for byte until the
    /// sheet's drawing is regenerated.
    /// </summary>
    public ChartCollection Charts { get; }

    /// <summary>Adds a chart with one series.</summary>
    /// <param name="kind">What to plot: column, bar, line, pie, doughnut or area.</param>
    /// <param name="anchor">The cells the chart covers, e.g. <c>"D2:K17"</c>.</param>
    /// <param name="valuesReference">Where the values come from, e.g. <c>"Данные!$B$2:$B$10"</c>.</param>
    /// <param name="categoriesReference">Where the category labels come from, or <see langword="null"/>.</param>
    /// <param name="title">The chart title, or <see langword="null"/>.</param>
    /// <param name="seriesName">The series name, or <see langword="null"/>.</param>
    public Chart AddChart(
        ChartKind kind, string anchor, string valuesReference,
        string? categoriesReference = null, string? title = null, string? seriesName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(valuesReference);
        return AddChart(kind, anchor, [new ChartSeriesDefinition(valuesReference, categoriesReference, seriesName)], title);
    }

    /// <summary>Adds a chart with any number of series.</summary>
    /// <param name="kind">What to plot: column, bar, line, pie, doughnut or area.</param>
    /// <param name="anchor">The cells the chart covers, e.g. <c>"D2:K17"</c>.</param>
    /// <param name="series">The series, in plot order.</param>
    /// <param name="title">The chart title, or <see langword="null"/>.</param>
    /// <remarks>
    /// The caches beside each reference are filled from the live cells at this moment, the way
    /// Excel fills them when it last drew the chart; a reference into cells that change later
    /// is refreshed by whatever next recalculates the chart.
    /// </remarks>
    public Chart AddChart(ChartKind kind, string anchor, IReadOnlyList<ChartSeriesDefinition> series, string? title = null)
    {
        ArgumentNullException.ThrowIfNull(series);
        if (series.Count == 0)
            throw new ArgumentException("A chart needs at least one series.", nameof(series));
        if (!ChartWriter.CanCreate(kind))
            throw new ArgumentException($"A {kind} chart cannot be created; only read.", nameof(kind));

        RangeRef cells = RangeRef.Parse(anchor);
        return Charts.Add(new Chart
        {
            RawXml = ChartWriter.Build(Workbook, kind, series, title),
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.TwoCell,
                From = cells.Start,
                To = new CellRef(cells.End.Row + 1, cells.End.Column + 1),
            },
            Name = $"Chart {Charts.Count + 1}",
        });
    }

    /// <summary>
    /// Threaded comments on cells of this sheet — the modern conversations, as opposed to the
    /// classic sticky notes in <see cref="Comments"/>.
    /// </summary>
    public ThreadedCommentCollection ThreadedComments { get; }

    /// <summary>Excel tables defined over ranges of this sheet.</summary>
    public TableCollection Tables { get; }

    /// <summary>Images placed on this sheet, back to front.</summary>
    public PictureCollection Pictures { get; } = new();

    /// <summary>
    /// Whether the drawing this sheet was loaded from held anything besides the pictures and
    /// charts the model can write — for example a shape or text box. Regenerating the part would
    /// lose those, so the saver warns rather than dropping them quietly.
    /// </summary>
    internal bool DrawingHasOtherShapes { get; set; }

    /// <summary>
    /// Range covered by the sheet's auto filter, or <see langword="null"/> when there is none.
    /// Filter criteria of a loaded file are kept as long as the range is not moved; the model
    /// does not represent them, so a new range starts with the dropdowns unfiltered.
    /// </summary>
    public RangeRef? AutoFilterRange
    {
        get => _autoFilterRange;
        set
        {
            if (_autoFilterRange != value)
                AutoFilterIsActive = false;
            _autoFilterRange = value;
        }
    }

    /// <summary>
    /// Whether the auto filter of the loaded file had criteria on at least one column, so that
    /// rows failing them are hidden. The file records those rows exactly like manually hidden
    /// ones, and this flag is what tells the two apart — SUBTOTAL depends on the difference.
    /// Moving the filter clears it, as the criteria no longer describe the rows.
    /// </summary>
    public bool AutoFilterIsActive { get; internal set; }

    /// <summary>
    /// Ranges of filtered tables. A table carries its own filter next to the sheet's, and its
    /// rejected rows are hidden the same way, so both have to be known.
    /// </summary>
    internal void MarkFilteredArea(RangeRef range) => (_filteredAreas ??= []).Add(range);

    /// <summary>
    /// Whether the row is hidden because a filter rejected it rather than because somebody
    /// hid it. Rows on a filter's header and past its range are never filtered away.
    /// </summary>
    internal bool IsFilteredAway(int row)
    {
        if (AutoFilterIsActive && AutoFilterRange is { } own && Covers(own, row))
            return true;

        if (_filteredAreas is null)
            return false;

        foreach (RangeRef area in _filteredAreas)
        {
            if (Covers(area, row))
                return true;
        }

        return false;

        static bool Covers(RangeRef range, int row) => row > range.Start.Row && row <= range.End.Row;
    }

    /// <summary>
    /// Where this sheet's part lived in the package it was loaded from, and the relationship
    /// id the workbook used for it. Saving reuses both so that preserved parts pointing at
    /// the sheet — and the sheet's own relative targets — keep resolving.
    /// </summary>
    internal string? SourcePartPath { get; set; }

    /// <inheritdoc cref="SourcePartPath"/>
    internal string? SourceRelationshipId { get; set; }

    /// <summary>
    /// The <c>sheetId</c> the loaded workbook gave this sheet, or 0 for a sheet added here.
    /// Revisions and extension lists point at sheets by this number, so renumbering would
    /// silently break them.
    /// </summary>
    internal int SourceSheetId { get; set; }

    /// <summary>Relationships the sheet part declared and the model does not regenerate.</summary>
    internal List<OpcRelationship> SourceRelationships { get; } = [];

    internal SheetGrid<CellValue> Values { get; } = new();

    internal SheetGrid<int> StyleIndexes { get; } = new();

    internal SheetGrid<string?> Formulas { get; } = new();

    internal Dictionary<int, ColumnProperties> ColumnTable { get; } = [];

    internal Dictionary<int, RowProperties> RowTable { get; } = [];

    internal List<RangeRef> Merges { get; } = [];

    /// <summary>
    /// Attributes of cells' formula elements the model does not represent, keyed by cell. Sparse
    /// by nature: almost no cell has any.
    /// </summary>
    internal Dictionary<CellRef, FormulaMarkup> FormulaMarkupTable { get; } = [];

    /// <summary>
    /// Attributes of cells the model does not represent — the <c>cm</c> and <c>vm</c> indexes
    /// that tie a cell to the metadata part behind rich values and dynamic arrays. The part
    /// itself is preserved by index, so carrying these keeps the link intact.
    /// </summary>
    internal Dictionary<CellRef, List<KeyValuePair<string, string>>> CellMarkupTable { get; } = [];

    /// <summary>
    /// Rich shared-string runs tied to their source cells. The visible text is retained beside
    /// the markup so a changed value is written as a new plain string instead of inheriting
    /// obsolete character formatting.
    /// </summary>
    internal Dictionary<CellRef, RichStringMarkup> RichStringTable { get; } = [];

    /// <summary>
    /// The spill ranges of the sheet's dynamic array formulas, anchor to rectangle. Filled by
    /// the calculation engine — the core stores the result but never computes one — and read
    /// back by the writer, which marks each anchor as a dynamic array in the file.
    /// </summary>
    internal Dictionary<CellRef, RangeRef> SpillTable { get; } = [];

    /// <summary>
    /// The anchor whose spill covers a cell, or null. The anchor covers itself.
    /// </summary>
    internal CellRef? SpillAnchorCovering(CellRef cell)
    {
        if (SpillTable.Count == 0)
            return null;

        foreach ((CellRef anchor, RangeRef range) in SpillTable)
        {
            if (range.Contains(cell))
                return anchor;
        }

        return null;
    }

    /// <summary>Content of the sheet part the model does not represent, replayed on save.</summary>
    internal WorksheetMarkup? Markup { get; set; }

    /// <summary>
    /// Shared and array formula groups read from the file. Shared groups disappear once the
    /// calculation engine expands them into per-cell formulas; array groups are preserved so
    /// the master keeps its <c>ref</c> attribute on save.
    /// </summary>
    internal List<FormulaGroup> FormulaGroups { get; } = [];

    /// <summary>
    /// Rewrites every formula and reference of this sheet that named <paramref name="oldName"/>.
    /// Called for each sheet of the workbook when any one of them is renamed.
    /// </summary>
    internal void RetargetSheetName(string oldName, string newName)
    {
        Formulas.Transform(formula => FormulaText.RenameSheet(formula, oldName, newName));

        foreach (DataValidation validation in Validations)
        {
            validation.Formula1 = FormulaText.RenameSheet(validation.Formula1, oldName, newName);
            validation.Formula2 = FormulaText.RenameSheet(validation.Formula2, oldName, newName);
        }

        foreach (ConditionalFormatRule rule in ConditionalFormats)
        {
            if (rule.Formulas.Count > 0)
                rule.Formulas = [.. rule.Formulas.Select(f => FormulaText.RenameSheet(f, oldName, newName)!)];
        }

        foreach (Hyperlink hyperlink in Hyperlinks)
        {
            if (!hyperlink.IsExternal)
                hyperlink.Location = FormulaText.RenameSheet(hyperlink.Location, oldName, newName);
        }

        foreach (Chart chart in Charts)
        {
            ReadOnlyMemory<byte> rewritten = ChartPart.RetargetSheetName(
                chart.RawXml, oldName, newName, out _, out bool changed);
            if (!changed)
                continue;

            chart.ReplaceRawXml(rewritten);
            if (chart.SourcePartPath is null || Workbook.Preserved?.Parts.ContainsKey(chart.SourcePartPath) != true)
                Charts.MarkModified();
        }
    }

    /// <summary>Notifies the workbook that a cell changed; <paramref name="structural"/> marks formula edits.</summary>
    internal void RaiseCellChanged(CellRef cell, bool structural)
    {
        Workbook.IsFullyCalculated = false;
        Workbook.CellChanged?.Invoke(this, cell, structural);
    }

    /// <summary>
    /// Notifies the workbook that a row's visibility changed. SUBTOTAL and AGGREGATE read it,
    /// so their results go stale even though no cell was touched.
    /// </summary>
    internal void RaiseVisibilityChanged()
    {
        Workbook.IsFullyCalculated = false;
        Workbook.VisibilityChanged?.Invoke(this);
    }

    /// <summary>Reads or writes a cell value by A1 reference: <c>sheet["B2"] = 42;</c>.</summary>
    public CellValue this[string a1]
    {
        get => Cell(a1).Value;
        set => Cell(a1).SetValue(value);
    }

    /// <summary>Reads or writes a cell value by 1-based row and column.</summary>
    public CellValue this[int row, int column]
    {
        get => Cell(row, column).Value;
        set => Cell(row, column).SetValue(value);
    }

    /// <summary>Returns a cell facade for an A1 reference.</summary>
    public Cell Cell(string a1) => new(this, CellRef.Parse(a1));

    /// <summary>Returns a cell facade for 1-based coordinates.</summary>
    public Cell Cell(int row, int column) => new(this, new CellRef(row, column));

    /// <summary>Returns a cell facade for a reference.</summary>
    public Cell Cell(CellRef reference) => new(this, reference);

    /// <summary>Returns a range facade for an A1 range like <c>A1:C5</c>.</summary>
    public CellRange Range(string a1) => new(this, RangeRef.Parse(a1));

    /// <summary>Returns a range facade.</summary>
    public CellRange Range(RangeRef reference) => new(this, reference);

    /// <summary>Merged ranges of the sheet.</summary>
    public IReadOnlyList<RangeRef> MergedRanges => Merges;

    /// <summary>Merges a range. Overlapping an existing merge throws.</summary>
    public void Merge(RangeRef range)
    {
        foreach (RangeRef existing in Merges)
        {
            if (existing.Intersects(range))
                throw new InvalidOperationException($"Range {range} overlaps the existing merge {existing}.");
        }

        Merges.Add(range);
    }

    /// <summary>Removes a merge previously added with <see cref="Merge"/>.</summary>
    public bool Unmerge(RangeRef range) => Merges.Remove(range);

    /// <summary>Puts filter dropdowns on the header row of a range.</summary>
    public void AutoFilter(RangeRef range) => AutoFilterRange = range;

    /// <summary>Puts filter dropdowns on the header row of an A1 range like <c>A1:D100</c>.</summary>
    public void AutoFilter(string a1) => AutoFilterRange = RangeRef.Parse(a1);

    /// <summary>Removes the sheet's auto filter.</summary>
    public void RemoveAutoFilter() => AutoFilterRange = null;

    /// <summary>
    /// Inserts blank rows above <paramref name="row"/>, pushing everything below it down. Cells,
    /// styles, merges, links, notes, rules and tables move with it, and every formula in the
    /// workbook is re-aimed. New rows inherit the height and visibility of the row above.
    /// </summary>
    /// <param name="row">1-based row the first new row takes the place of.</param>
    /// <param name="count">How many rows to insert.</param>
    public void InsertRows(int row, int count = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        WorksheetEditor.Shift(this, rows: true, row, count);
    }

    /// <summary>
    /// Deletes rows and pulls the rows below them up. Formulas that pointed into the deleted
    /// rows become <c>#REF!</c>, as they do in Excel.
    /// </summary>
    /// <param name="row">1-based first row to delete.</param>
    /// <param name="count">How many rows to delete.</param>
    public void DeleteRows(int row, int count = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        WorksheetEditor.Shift(this, rows: true, row, -count);
    }

    /// <summary>Column-wise counterpart of <see cref="InsertRows"/>.</summary>
    /// <param name="column">1-based column the first new column takes the place of.</param>
    /// <param name="count">How many columns to insert.</param>
    public void InsertColumns(int column, int count = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        WorksheetEditor.Shift(this, rows: false, column, count);
    }

    /// <summary>Column-wise counterpart of <see cref="DeleteRows"/>.</summary>
    /// <param name="column">1-based first column to delete.</param>
    /// <param name="count">How many columns to delete.</param>
    public void DeleteColumns(int column, int count = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        WorksheetEditor.Shift(this, rows: false, column, -count);
    }

    /// <summary>Returns a facade for one row: height, visibility, outline level.</summary>
    public SheetRow Row(int number) => new(this, number);

    /// <summary>Returns a facade for a band of rows, for grouping and bulk visibility.</summary>
    public SheetRowRange Rows(int first, int last) => new(this, first, last);

    /// <summary>Returns a facade for one column: width, visibility, outline level.</summary>
    public SheetColumn Column(int number) => new(this, number);

    /// <summary>Returns a facade for one column by letter name, e.g. <c>"B"</c>.</summary>
    public SheetColumn Column(string name) => new(this, ColumnName.Parse(name));

    /// <summary>Returns a facade for a band of columns, for grouping and bulk visibility.</summary>
    public SheetColumnRange Columns(int first, int last) => new(this, first, last);

    /// <summary>Rows that carry a height, visibility or outline setting, keyed by row number.</summary>
    public IReadOnlyDictionary<int, RowProperties> ConfiguredRows => RowTable;

    /// <summary>Columns that carry a width, visibility or outline setting, keyed by column number.</summary>
    public IReadOnlyDictionary<int, ColumnProperties> ConfiguredColumns => ColumnTable;

    /// <summary>Sets a custom column width in characters.</summary>
    public void SetColumnWidth(int column, double width)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        Column(column).Width = width;
    }

    /// <summary>Sets a custom row height in points.</summary>
    public void SetRowHeight(int row, double height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        Row(row).Height = height;
    }

    /// <summary>Formats the worksheet's used range as a CSV or TSV string.</summary>
    public string ToCsv(CsvOptions? options = null) =>
        CsvWriter.ExportWorksheet(this, options ?? CsvOptions.Default);

    /// <summary>Writes the worksheet's used range to a stream as CSV or TSV.</summary>
    public async ValueTask SaveCsvAsync(
        Stream stream, CsvOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        CsvOptions effective = options ?? CsvOptions.Default;
        string csv = CsvWriter.ExportWorksheet(this, effective);
        byte[] bytes = effective.Encoding.GetBytes(csv);
        await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
        if (!effective.LeaveOpen)
            await stream.DisposeAsync().ConfigureAwait(false);
    }

    /// <summary>Writes the worksheet's used range to a file as CSV or TSV, overwriting any existing file.</summary>
    public async ValueTask SaveCsvAsync(
        string path, CsvOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var stream = new FileStream(
            path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 64 * 1024, useAsync: true);
        await using (stream.ConfigureAwait(false))
        {
            await SaveCsvAsync(stream, (options ?? CsvOptions.Default) with { LeaveOpen = true }, cancellationToken)
                .ConfigureAwait(false);
        }
    }

    /// <summary>
    /// The rectangle spanned by everything the sheet stores (values, styles or formulas),
    /// or <see langword="null"/> for an empty sheet. Upper bound: clearing cells does not shrink it.
    /// </summary>
    public RangeRef? UsedRange
    {
        get
        {
            int maxRow = Math.Max(Values.MaxRow, Math.Max(StyleIndexes.MaxRow, Formulas.MaxRow));
            int maxColumn = Math.Max(Values.MaxColumn, Math.Max(StyleIndexes.MaxColumn, Formulas.MaxColumn));
            if (maxRow == 0 || maxColumn == 0)
                return null;
            return new RangeRef(new CellRef(1, 1), new CellRef(maxRow, maxColumn));
        }
    }
}
