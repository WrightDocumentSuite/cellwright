namespace Cellwright;

/// <summary>Which way round the paper goes.</summary>
public enum PageOrientation
{
    /// <summary>Whatever the printer decides; the attribute is left out.</summary>
    Default,

    /// <summary>Taller than wide.</summary>
    Portrait,

    /// <summary>Wider than tall.</summary>
    Landscape,
}

/// <summary>The direction in which a sheet wider and taller than one page is paginated.</summary>
public enum PageOrder
{
    /// <summary>All the way down the first column of pages, then across.</summary>
    DownThenOver,

    /// <summary>All the way across the first row of pages, then down.</summary>
    OverThenDown,
}

/// <summary>Printing margins, in inches — the unit the file format uses.</summary>
/// <param name="Left">Left margin.</param>
/// <param name="Right">Right margin.</param>
/// <param name="Top">Top margin, from the paper edge to the body.</param>
/// <param name="Bottom">Bottom margin, from the body to the paper edge.</param>
/// <param name="Header">Distance from the paper edge to the header.</param>
/// <param name="Footer">Distance from the footer to the paper edge.</param>
public readonly record struct PageMargins(
    double Left, double Right, double Top, double Bottom, double Header, double Footer)
{
    /// <summary>Excel's "Normal" margins.</summary>
    public static PageMargins Normal { get; } = new(0.7, 0.7, 0.75, 0.75, 0.3, 0.3);

    /// <summary>Excel's "Wide" margins.</summary>
    public static PageMargins Wide { get; } = new(1, 1, 1, 1, 0.5, 0.5);

    /// <summary>Excel's "Narrow" margins.</summary>
    public static PageMargins Narrow { get; } = new(0.25, 0.25, 0.75, 0.75, 0.3, 0.3);
}

/// <summary>
/// The three lines of a header or footer. Each holds the codes Excel uses: <c>&amp;L</c>,
/// <c>&amp;C</c> and <c>&amp;R</c> split the line into sections, <c>&amp;P</c> is the page
/// number, <c>&amp;N</c> the page count, <c>&amp;A</c> the sheet name and <c>&amp;D</c> the date.
/// </summary>
public sealed record HeaderFooterText
{
    /// <summary>Text of every page, or of the odd pages when the others differ.</summary>
    public string? Default { get; init; }

    /// <summary>Text of the even pages, used only when <see cref="PageSetup.DifferentOddEven"/> is set.</summary>
    public string? Even { get; init; }

    /// <summary>Text of the first page, used only when <see cref="PageSetup.DifferentFirstPage"/> is set.</summary>
    public string? First { get; init; }

    /// <summary>Whether every line is empty.</summary>
    public bool IsEmpty => string.IsNullOrEmpty(Default) && string.IsNullOrEmpty(Even) && string.IsNullOrEmpty(First);
}

/// <summary>
/// How a worksheet is printed: paper, orientation, scaling, margins, headers and page breaks.
/// Reading a file fills this in; the first change makes the model the source of truth and the
/// original markup is regenerated from it.
/// </summary>
/// <example>
/// <code>
/// sheet.PageSetup.Orientation = PageOrientation.Landscape;
/// sheet.PageSetup.FitToWidth = 1;
/// sheet.PageSetup.Header = new HeaderFooterText { Default = "&amp;CQuarterly report" };
/// sheet.SetPrintArea("A1:H40");
/// sheet.SetRepeatRows(1, 1);
/// </code>
/// </example>
public sealed class PageSetup
{
    private PageOrientation _orientation;
    private int _paperSize;
    private int _scale = 100;
    private int? _fitToWidth;
    private int? _fitToHeight;
    private PageMargins? _margins;
    private HeaderFooterText? _header;
    private HeaderFooterText? _footer;
    private bool _printGridLines;
    private bool _printHeadings;
    private bool _centerHorizontally;
    private bool _centerVertically;
    private bool _blackAndWhite;
    private bool _draft;
    private int _firstPageNumber;
    private PageOrder _order;
    private bool _differentOddEven;
    private bool _differentFirstPage;
    private bool _scaleWithDocument = true;
    private bool _alignWithMargins = true;
    private bool _autoPageBreaks = true;
    private List<int>? _rowBreaks;
    private List<int>? _columnBreaks;

    /// <summary>Whether the model, rather than the loaded file, owns the print settings.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>Which way round the paper goes.</summary>
    public PageOrientation Orientation
    {
        get => _orientation;
        set => Set(ref _orientation, value);
    }

    /// <summary>Paper size code; 1 is Letter, 8 is A3, 9 is A4. Zero leaves the printer's default.</summary>
    public int PaperSize
    {
        get => _paperSize;
        set => Set(ref _paperSize, value);
    }

    /// <summary>Print scale in percent. Ignored when <see cref="FitToWidth"/> or <see cref="FitToHeight"/> is set.</summary>
    public int Scale
    {
        get => _scale;
        set => Set(ref _scale, value);
    }

    /// <summary>How many pages wide the sheet is squeezed into, or <see langword="null"/> to scale freely.</summary>
    public int? FitToWidth
    {
        get => _fitToWidth;
        set => Set(ref _fitToWidth, value);
    }

    /// <summary>How many pages tall the sheet is squeezed into, or <see langword="null"/> to scale freely.</summary>
    public int? FitToHeight
    {
        get => _fitToHeight;
        set => Set(ref _fitToHeight, value);
    }

    /// <summary>Whether the sheet is fitted to a page count rather than scaled by percent.</summary>
    public bool FitToPage => _fitToWidth is not null || _fitToHeight is not null;

    /// <summary>Whether Excel may add page breaks of its own on top of the manual ones.</summary>
    public bool AutoPageBreaks
    {
        get => _autoPageBreaks;
        set => Set(ref _autoPageBreaks, value);
    }

    /// <summary>
    /// Whether the loaded file had the fit-to-page flag. The page counts sit on another element
    /// and mean nothing without it, so reading them waits for this.
    /// </summary>
    internal bool FitToPageRequested { get; set; }

    /// <summary>Margins in inches, or <see langword="null"/> while the file's own are kept.</summary>
    public PageMargins? Margins
    {
        get => _margins;
        set => Set(ref _margins, value);
    }

    /// <summary>Header lines.</summary>
    public HeaderFooterText? Header
    {
        get => _header;
        set => Set(ref _header, value);
    }

    /// <summary>Footer lines.</summary>
    public HeaderFooterText? Footer
    {
        get => _footer;
        set => Set(ref _footer, value);
    }

    /// <summary>Whether the grid is printed.</summary>
    public bool PrintGridLines
    {
        get => _printGridLines;
        set => Set(ref _printGridLines, value);
    }

    /// <summary>Whether the row numbers and column letters are printed.</summary>
    public bool PrintHeadings
    {
        get => _printHeadings;
        set => Set(ref _printHeadings, value);
    }

    /// <summary>Whether the content is centered between the left and right margins.</summary>
    public bool CenterHorizontally
    {
        get => _centerHorizontally;
        set => Set(ref _centerHorizontally, value);
    }

    /// <summary>Whether the content is centered between the top and bottom margins.</summary>
    public bool CenterVertically
    {
        get => _centerVertically;
        set => Set(ref _centerVertically, value);
    }

    /// <summary>Whether colors are dropped when printing.</summary>
    public bool BlackAndWhite
    {
        get => _blackAndWhite;
        set => Set(ref _blackAndWhite, value);
    }

    /// <summary>Whether the sheet prints without graphics, for a quick proof.</summary>
    public bool Draft
    {
        get => _draft;
        set => Set(ref _draft, value);
    }

    /// <summary>Number the first page carries, or 0 to keep Excel's automatic numbering.</summary>
    public int FirstPageNumber
    {
        get => _firstPageNumber;
        set => Set(ref _firstPageNumber, value);
    }

    /// <summary>Order in which pages are numbered.</summary>
    public PageOrder Order
    {
        get => _order;
        set => Set(ref _order, value);
    }

    /// <summary>Whether even pages get their own header and footer.</summary>
    public bool DifferentOddEven
    {
        get => _differentOddEven;
        set => Set(ref _differentOddEven, value);
    }

    /// <summary>Whether the first page gets its own header and footer.</summary>
    public bool DifferentFirstPage
    {
        get => _differentFirstPage;
        set => Set(ref _differentFirstPage, value);
    }

    /// <summary>Whether headers and footers follow the print scale.</summary>
    public bool ScaleHeaderWithDocument
    {
        get => _scaleWithDocument;
        set => Set(ref _scaleWithDocument, value);
    }

    /// <summary>Whether headers and footers line up with the page margins.</summary>
    public bool AlignHeaderWithMargins
    {
        get => _alignWithMargins;
        set => Set(ref _alignWithMargins, value);
    }

    /// <summary>Rows a page break starts above, 1-based and ascending.</summary>
    public IReadOnlyList<int> RowBreaks => _rowBreaks ?? (IReadOnlyList<int>)[];

    /// <summary>Columns a page break starts left of, 1-based and ascending.</summary>
    public IReadOnlyList<int> ColumnBreaks => _columnBreaks ?? (IReadOnlyList<int>)[];

    /// <summary>Puts a horizontal page break above the given row.</summary>
    public void AddRowBreak(int row)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(row, 2);
        AddBreak(ref _rowBreaks, row);
    }

    /// <summary>Puts a vertical page break to the left of the given column.</summary>
    public void AddColumnBreak(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 2);
        AddBreak(ref _columnBreaks, column);
    }

    /// <summary>Removes every page break.</summary>
    public void ClearBreaks()
    {
        if (_rowBreaks is null && _columnBreaks is null)
            return;

        _rowBreaks = null;
        _columnBreaks = null;
        IsModified = true;
    }

    /// <summary>Marks the settings as owned by the model, e.g. after a structural edit moved a break.</summary>
    internal void MarkModified() => IsModified = true;

    /// <summary>Records a value read from a file without making the model the source of truth.</summary>
    internal void SetFromFile(Action<PageSetup> apply)
    {
        bool modified = IsModified;
        apply(this);
        IsModified = modified;
    }

    /// <summary>Replaces the breaks read from a file, keeping ownership with the file.</summary>
    internal void SetBreaksFromFile(List<int>? rows, List<int>? columns)
    {
        _rowBreaks = rows;
        _columnBreaks = columns;
    }

    /// <summary>Replaces the breaks of one axis after a structural edit moved them.</summary>
    internal void ReplaceBreaks(List<int>? rows, List<int>? columns)
    {
        if (rows is not null)
            _rowBreaks = rows.Count == 0 ? null : rows;
        if (columns is not null)
            _columnBreaks = columns.Count == 0 ? null : columns;
        IsModified = true;
    }

    private void AddBreak(ref List<int>? breaks, int at)
    {
        breaks ??= [];
        int position = breaks.BinarySearch(at);
        if (position >= 0)
            return;

        breaks.Insert(~position, at);
        IsModified = true;
    }

    private void Set<T>(ref T field, T value)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return;

        field = value;
        IsModified = true;
    }
}
