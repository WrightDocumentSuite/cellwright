using Cellwright.IO;
using Cellwright.Markdown;
using Cellwright.Model;
using Cellwright.Styles;

namespace Cellwright;

/// <summary>
/// In-memory workbook document. Purely a data model: loading pulls everything out of the
/// package and closes it; saving streams the model through <see cref="XlsxWriter"/>.
/// Cells live in data-oriented column slices — there are no per-cell objects to collect.
/// </summary>
public sealed class Workbook
{
    private Workbook()
    {
        Sheets = new SheetCollection(this);
    }

    /// <summary>The worksheets of the workbook.</summary>
    public SheetCollection Sheets { get; }

    /// <summary>Named references and constants of the workbook.</summary>
    public DefinedNameCollection Names { get; } = new();

    /// <summary>
    /// The people who wrote the workbook's threaded comments. Workbook-wide because a person who
    /// comments on three sheets is one person; a comment names its author from here.
    /// </summary>
    public PersonCollection People { get; } = new();

    /// <summary>Date system used for serial conversions when saving.</summary>
    public ExcelDateSystem DateSystem { get; set; }

    /// <summary>Protection of the sheet list and the window. Off until <c>Protect</c> is called.</summary>
    public WorkbookProtection Protection { get; } = new();

    /// <summary>
    /// What the file says about itself: title, author, company, custom fields. Saving always
    /// writes them, so a workbook created here arrives with the same document properties
    /// Excel's own files carry.
    /// </summary>
    public WorkbookProperties Properties { get; } = new();

    /// <summary>Warnings and repairs from the load; <see langword="null"/> for created workbooks.</summary>
    public LoadDiagnostics? LoadDiagnostics { get; private set; }

    /// <summary>
    /// The colour scheme a themed <see cref="ExcelColor"/> names a slot of. A file that carries
    /// no theme, and a workbook created here, use the one Excel starts with.
    /// </summary>
    public WorkbookTheme Theme { get; internal set; } = WorkbookTheme.Office;

    /// <summary>
    /// Called when an edit has to give something up — a structural change dropping markup whose
    /// ranges it cannot follow, say. Load-time warnings go to <see cref="LoadDiagnostics"/>.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; set; }

    internal StylePool StylePool { get; } = new();

    /// <summary>
    /// The differential formats of the loaded file, parsed by position. A conditional rule
    /// shows the one it points at through <see cref="ConditionalFormatRule.Style"/>; the saver
    /// compares against this list to keep an untouched rule on its original <c>dxf</c> index.
    /// </summary>
    internal IReadOnlyList<DifferentialStyle> LoadedDifferentialStyles { get; set; } = [];

    /// <summary>
    /// The workbooks this one links to, with the cell values cached at the last refresh
    /// (§18.14). Formulas like <c>[1]Sheet1!A1</c> answer from these caches; position in the
    /// list is the number in the brackets, counted from one.
    /// </summary>
    public IReadOnlyList<ExternalBook> ExternalBooks { get; internal set; } = [];

    /// <summary>
    /// Content of the loaded package that the model does not represent — charts, pivot
    /// tables, images, macros, printer settings. Saving copies it through unchanged.
    /// <see langword="null"/> for workbooks created from scratch or loaded with
    /// <see cref="LoadOptions.PreserveUnknownParts"/> turned off.
    /// </summary>
    internal PreservedPackage? Preserved { get; set; }

    /// <summary>
    /// Raised when a cell's value, formula or content is changed through the model.
    /// The calculation engine subscribes to keep its dirty set current; when nothing is
    /// subscribed the cost is a null check on the mutation path.
    /// </summary>
    internal Action<Worksheet, CellRef, bool>? CellChanged { get; set; }

    /// <summary>
    /// Raised when a row is hidden or shown. Only the aggregates that skip hidden rows care,
    /// so the engine keeps a short list of those instead of dirtying the sheet.
    /// </summary>
    internal Action<Worksheet>? VisibilityChanged { get; set; }

    /// <summary>
    /// Whether every formula holds an up-to-date computed value. Set by the calculation
    /// engine, cleared by any model mutation; controls the <c>fullCalcOnLoad</c> flag on save.
    /// </summary>
    internal bool IsFullyCalculated { get; set; }

    internal int NextSheetId { get; set; }

    /// <summary>Creates an empty workbook.</summary>
    public static Workbook Create() => new();

    /// <summary>
    /// Finds a table by name and the sheet holding it. Table names are unique across a workbook,
    /// which is what lets a structured reference name one without saying where it lives.
    /// </summary>
    internal (Worksheet Sheet, ExcelTable Table)? LocateTable(string name)
    {
        foreach (Worksheet sheet in Sheets)
        {
            if (sheet.Tables.Find(name) is { } table)
                return (sheet, table);
        }

        return null;
    }

    /// <summary>
    /// Points everything that named a sheet at its new name: the formulas of every sheet, the
    /// defined names, and the chart and pivot definitions of a preserved package.
    /// </summary>
    internal void RetargetSheetName(string oldName, string newName)
    {
        foreach (Worksheet sheet in Sheets)
            sheet.RetargetSheetName(oldName, newName);

        foreach (DefinedName name in Names)
            name.RefersTo = FormulaText.RenameSheet(name.RefersTo, oldName, newName) ?? name.RefersTo;

        Preserved?.RetargetSheetName(oldName, newName);
    }

    /// <summary>
    /// Interns a style and returns its handle for use with <see cref="Cell.SetStyle(StyleId)"/>.
    /// Equal styles share one handle.
    /// </summary>
    public StyleId RegisterStyle(CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new StyleId(StylePool.Intern(style));
    }

    /// <summary>Returns the style record behind a handle.</summary>
    public CellStyle GetStyle(StyleId id) =>
        (uint)id.Index < (uint)StylePool.Styles.Count ? StylePool.Styles[id.Index] : CellStyle.Default;

    /// <summary>Loads a workbook from a stream.</summary>
    public static async ValueTask<Workbook> LoadAsync(
        Stream stream, LoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        LoadOptions effective = options ?? LoadOptions.Default;
        var workbook = new Workbook();
        await WorkbookLoader.LoadAsync(workbook, stream, effective, cancellationToken).ConfigureAwait(false);
        return workbook;
    }

    /// <summary>Loads a workbook from a file.</summary>
    public static ValueTask<Workbook> LoadAsync(
        string path, LoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 64 * 1024, useAsync: true);
        return LoadWithOwnedStreamAsync(stream, options, cancellationToken);
    }

    /// <summary>Saves the workbook to a stream.</summary>
    public async ValueTask SaveAsync(
        Stream stream, SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        await WorkbookSaver.SaveAsync(this, stream, options ?? SaveOptions.Default, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the workbook to a file, overwriting any existing one.</summary>
    public async ValueTask SaveAsync(
        string path, SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var stream = new FileStream(
            path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 64 * 1024, useAsync: true);
        await using (stream.ConfigureAwait(false))
        {
            await SaveAsync(stream, (options ?? SaveOptions.Default) with { LeaveOpen = true }, cancellationToken)
                .ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Renders the workbook as GitHub Flavored Markdown: a section per visible sheet, its used
    /// range as a table with the values shown the way Excel shows them, and its pictures as a
    /// gallery beneath. A pure conversion — nothing is written anywhere, and the returned
    /// document carries the text and the image files side by side.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <see cref="MarkdownExportOptions.MediaDirectoryName"/> is not a safe relative directory.
    /// </exception>
    public MarkdownDocument ToMarkdown(MarkdownExportOptions? options = null) =>
        MarkdownExporter.Export(this, options ?? MarkdownExportOptions.Default);

    /// <summary>
    /// Writes the workbook into a directory as <see cref="MarkdownDocument.DefaultFileName"/>
    /// beside the media directory its image links point into, creating what is not there.
    /// </summary>
    public ValueTask ExportMarkdownAsync(
        string directoryPath, MarkdownExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(directoryPath);
        return ToMarkdown(options).SaveAsync(directoryPath, cancellationToken);
    }

    internal void AttachDiagnostics(LoadDiagnostics diagnostics) => LoadDiagnostics = diagnostics;

    internal void SetDateSystem(ExcelDateSystem system) => DateSystem = system;

    private static async ValueTask<Workbook> LoadWithOwnedStreamAsync(
        FileStream stream, LoadOptions? options, CancellationToken cancellationToken)
    {
        await using (stream.ConfigureAwait(false))
        {
            return await LoadAsync(stream, (options ?? LoadOptions.Default) with { LeaveOpen = true }, cancellationToken)
                .ConfigureAwait(false);
        }
    }
}
