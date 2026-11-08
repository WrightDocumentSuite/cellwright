using Cellwright.IO;

namespace Cellwright.Streaming;

/// <summary>
/// Receiver for the full content of one worksheet part. Keeps the streaming reader
/// unaware of the DOM: the loader implements this to fill a <see cref="Worksheet"/>.
/// </summary>
internal interface ISheetContentSink
{
    /// <summary>A frozen pane declaration.</summary>
    public void OnFrozenPane(int rows, int columns);

    /// <summary>The sheet's tab color.</summary>
    public void OnTabColor(ExcelColor color);

    /// <summary>The sheet view suppresses grid lines.</summary>
    public void OnGridLinesHidden();

    /// <summary>The sheet view's zoom percentage.</summary>
    public void OnZoomScale(int percent);

    /// <summary>A column definition: width, visibility, outline grouping and inherited style.</summary>
    public void OnColumn(in ColumnDefinition definition);

    /// <summary>A row definition: height, visibility, outline grouping and inherited style.</summary>
    public void OnRow(in RowDefinition definition);

    /// <summary>One cell with its raw style-part index and optional formula text.</summary>
    public void OnCell(CellRef cell, in CellValue value, int rawStyleIndex, string? formula);

    /// <summary>
    /// Rich-text markup of a shared-string cell, together with the plain text used to detect
    /// whether the value was edited before it is saved again.
    /// </summary>
    public void OnRichString(CellRef cell, string text, string markup);

    /// <summary>One merged range.</summary>
    public void OnMerge(RangeRef range);

    /// <summary>
    /// The range the sheet's auto filter covers, and whether any column carries criteria —
    /// the only clue a file gives that its hidden rows were hidden by filtering.
    /// </summary>
    public void OnAutoFilter(RangeRef range, bool hasCriteria);

    /// <summary>One data validation rule.</summary>
    public void OnDataValidation(DataValidation validation);

    /// <summary>One conditional formatting rule.</summary>
    public void OnConditionalFormat(ConditionalFormatRule rule);

    /// <summary>The sheet's print settings, to be filled in from a captured element.</summary>
    public PageSetup PageSetup { get; }

    /// <summary>The sheet's protection, to be filled in from a captured element.</summary>
    public SheetProtection Protection { get; }

    /// <summary>
    /// A hyperlink over a range. External targets arrive as a relationship id to resolve
    /// against the sheet's relationships; internal ones carry a <paramref name="location"/>.
    /// </summary>
    public void OnHyperlink(RangeRef range, string? relationshipId, string? location, string? tooltip, string? display);

    /// <summary>
    /// A shared or array formula master. <paramref name="sharedIndex"/> is the <c>si</c>
    /// attribute for shared groups, letting members be matched to this master.
    /// </summary>
    public void OnFormulaGroup(FormulaGroup group, int? sharedIndex);

    /// <summary>A cell that repeats a shared formula identified by <paramref name="sharedIndex"/>.</summary>
    public void OnSharedFormulaMember(CellRef cell, int sharedIndex);

    /// <summary>Attributes of a cell's formula element the model does not represent.</summary>
    public void OnFormulaMarkup(CellRef cell, FormulaMarkup markup);

    /// <summary>
    /// Attributes of a cell the model does not represent, chiefly the <c>cm</c> and <c>vm</c>
    /// indexes tying it to the metadata part.
    /// </summary>
    public void OnCellMarkup(CellRef cell, List<KeyValuePair<string, string>> attributes);

    /// <summary>Everything in the part the model does not represent, to be replayed on save.</summary>
    public void OnMarkup(WorksheetMarkup markup);
}
