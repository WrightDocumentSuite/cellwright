using Cellwright.IO;

namespace Cellwright;

/// <summary>
/// Options controlling how <see cref="XlsxWriter"/> produces a workbook.
/// </summary>
public sealed record XlsxWriterOptions
{
    internal static readonly XlsxWriterOptions Default = new();

    /// <summary>
    /// When <see langword="true"/>, text is written as inline strings instead of the shared-string
    /// table. Useful for huge streams of mostly-unique text: nothing accumulates in memory.
    /// Shared strings (the default) produce smaller files when text repeats.
    /// </summary>
    public bool UseInlineStrings { get; init; }

    /// <summary>Date system recorded in the workbook and used for serial conversions.</summary>
    public ExcelDateSystem DateSystem { get; init; } = ExcelDateSystem.Date1900;

    /// <summary>Keep the destination stream open after the writer is disposed. Ignored for file paths.</summary>
    public bool LeaveOpen { get; init; }

    /// <summary>
    /// Ask Excel to recalculate every formula when the file is opened. Enabled by default
    /// because values written next to formulas are only as fresh as the caller made them;
    /// the calculation engine turns it off once it has computed the whole workbook.
    /// </summary>
    public bool FullCalculationOnLoad { get; init; } = true;
}

/// <summary>
/// Per-worksheet options supplied to <see cref="XlsxWriter.BeginSheetAsync"/>.
/// </summary>
public sealed record SheetOptions
{
    internal static readonly SheetOptions Default = new();

    /// <summary>Number of top rows frozen in place when scrolling.</summary>
    public int FrozenRows { get; init; }

    /// <summary>Number of left columns frozen in place when scrolling.</summary>
    public int FrozenColumns { get; init; }

    /// <summary>Column widths, visibility and outline levels, keyed by 1-based column number.</summary>
    public IReadOnlyDictionary<int, ColumnProperties>? Columns { get; init; }

    /// <summary>Row heights, visibility and outline levels, keyed by 1-based row number.</summary>
    public IReadOnlyDictionary<int, RowProperties>? Rows { get; init; }

    /// <summary>Color of the sheet tab; <see cref="ExcelColor.Auto"/> for the default.</summary>
    public ExcelColor TabColor { get; init; } = ExcelColor.Auto;

    /// <summary>Whether grid lines are drawn.</summary>
    public bool ShowGridLines { get; init; } = true;

    /// <summary>Zoom percentage of the sheet view.</summary>
    public int ZoomScale { get; init; } = 100;

    /// <summary>Hyperlinks to place over cells, already resolved to relationship ids where external.</summary>
    public IReadOnlyList<WrittenHyperlink>? Hyperlinks { get; init; }

    /// <summary>Range to put filter dropdowns on, or <see langword="null"/> for no filter.</summary>
    public RangeRef? AutoFilterRange { get; init; }

    /// <summary>Input rules to attach to ranges of the sheet.</summary>
    public IReadOnlyList<DataValidation>? Validations { get; init; }

    /// <summary>
    /// Whether the validations came from the model rather than a loaded file. When they did
    /// not, the source markup is replayed instead, keeping extensions the model cannot express.
    /// </summary>
    internal bool ValidationsChanged { get; init; }

    /// <summary>Conditional formatting rules to write, in the order they are tried.</summary>
    public IReadOnlyList<ConditionalFormatRule>? ConditionalFormats { get; init; }

    /// <summary>
    /// Differential format index per entry of <see cref="ConditionalFormats"/>, or -1 for a
    /// rule that paints nothing of its own. Resolved before the sheet opens, because the
    /// formats live in a part written after every sheet.
    /// </summary>
    internal IReadOnlyList<int>? ConditionalStyleIds { get; init; }

    /// <summary>
    /// Whether the rules came from the model rather than a loaded file. When they did not, the
    /// original blocks are replayed, keeping the x14 extensions of modern rule types.
    /// </summary>
    internal bool ConditionalFormatsChanged { get; init; }

    /// <summary>
    /// Print settings of the sheet. Written from the model only once they have been touched;
    /// until then the source elements are replayed with everything the model does not model.
    /// </summary>
    internal PageSetup? PageSetup { get; init; }

    /// <summary>Protection of the sheet; always generated, since the model covers the element fully.</summary>
    internal SheetProtection? Protection { get; init; }

    /// <summary>
    /// Ranges of array formulas, keyed by the master cell holding the formula text.
    /// Masters listed here are written with <c>t="array"</c> and their <c>ref</c>.
    /// </summary>
    public IReadOnlyDictionary<CellRef, RangeRef>? ArrayFormulaRanges { get; init; }

    /// <summary>Attributes of cells' formula elements the model does not represent, keyed by cell.</summary>
    internal IReadOnlyDictionary<CellRef, FormulaMarkup>? FormulaMarkup { get; init; }

    /// <summary>Attributes of cells the model does not represent, keyed by cell.</summary>
    internal IReadOnlyDictionary<CellRef, List<KeyValuePair<string, string>>>? CellMarkup { get; init; }

    /// <summary>Rich shared-string runs keyed by the source cell that owns them.</summary>
    internal IReadOnlyDictionary<CellRef, RichStringMarkup>? RichStrings { get; init; }

    /// <summary>How visible the sheet is in the Excel UI. At least one sheet must stay visible.</summary>
    public SheetVisibility Visibility { get; init; }

    /// <summary>
    /// Where to place the worksheet part. Saving a loaded workbook keeps each sheet at its
    /// original location so that relative targets of preserved parts keep resolving.
    /// </summary>
    internal string? PartPath { get; init; }

    /// <summary>Relationship id the workbook part should use for this sheet.</summary>
    internal string? RelationshipId { get; init; }

    /// <summary>Identifier the sheet had in the loaded workbook, or 0 to have one assigned.</summary>
    internal int SheetId { get; init; }

    /// <summary>Relationships the worksheet part declares: hyperlinks, comments, drawings, tables.</summary>
    internal IReadOnlyList<OpcRelationship>? Relationships { get; init; }

    /// <summary>Content of the source worksheet part that the model does not represent.</summary>
    internal WorksheetMarkup? Markup { get; init; }

    /// <summary>
    /// Whether the sheet's notes came from the model rather than a loaded file. When they did
    /// not, the preserved <c>legacyDrawing</c> element and its VML are replayed untouched.
    /// </summary>
    internal bool CommentsChanged { get; init; }

    /// <summary>Relationship of the regenerated VML drawing holding the note boxes.</summary>
    internal string? LegacyDrawingRelationshipId { get; init; }

    /// <summary>
    /// Whether the sheet's pictures came from the model rather than a loaded file. When they
    /// did not, the preserved <c>drawing</c> element and its part are replayed untouched.
    /// </summary>
    internal bool PicturesChanged { get; init; }

    /// <summary>Relationship of the regenerated drawing holding the pictures.</summary>
    internal string? DrawingRelationshipId { get; init; }

    /// <summary>
    /// Whether the sheet's tables came from the model rather than a loaded file. When they
    /// did not, the preserved <c>tableParts</c> element and the table parts are kept as is.
    /// </summary>
    internal bool TablesChanged { get; init; }

    /// <summary>Relationship ids of the regenerated table parts, in table order.</summary>
    internal IReadOnlyList<string>? TableRelationshipIds { get; init; }
}

/// <summary>
/// One hyperlink as the writer needs it: the relationship id is already allocated, because
/// the sheet part and its relationships are emitted in one pass.
/// </summary>
/// <param name="Range">Cells the link covers.</param>
/// <param name="RelationshipId">Relationship id for external links, otherwise <see langword="null"/>.</param>
/// <param name="Location">Location inside the target, or the target cell for internal links.</param>
/// <param name="Tooltip">Hover text.</param>
/// <param name="Display">Display text recorded alongside the link.</param>
public readonly record struct WrittenHyperlink(
    RangeRef Range, string? RelationshipId, string? Location, string? Tooltip, string? Display);
