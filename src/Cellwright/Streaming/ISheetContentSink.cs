using Cellwright.IO;

namespace Cellwright.Streaming;

/// <summary>
/// Receiver for the full content of one worksheet part. Keeps the streaming reader
/// unaware of the DOM: the loader implements this to fill a <see cref="Worksheet"/>.
/// </summary>
internal interface ISheetContentSink
{
    /// <summary>A frozen pane declaration.</summary>
    void OnFrozenPane(int rows, int columns);

    /// <summary>The sheet's tab color.</summary>
    void OnTabColor(ExcelColor color);

    /// <summary>The sheet view suppresses grid lines.</summary>
    void OnGridLinesHidden();

    /// <summary>The sheet view's zoom percentage.</summary>
    void OnZoomScale(int percent);

    /// <summary>A column definition: width, visibility, outline grouping and inherited style.</summary>
    void OnColumn(in ColumnDefinition definition);

    /// <summary>A row definition: height, visibility, outline grouping and inherited style.</summary>
    void OnRow(in RowDefinition definition);

    /// <summary>One cell with its raw style-part index and optional formula text.</summary>
    void OnCell(CellRef cell, in CellValue value, int rawStyleIndex, string? formula);

    /// <summary>
    /// Rich-text markup of a shared-string cell, together with the plain text used to detect
    /// whether the value was edited before it is saved again.
    /// </summary>
    void OnRichString(CellRef cell, string text, string markup);

    /// <summary>One merged range.</summary>
    void OnMerge(RangeRef range);

    /// <summary>
    /// The range the sheet's auto filter covers, and whether any column carries criteria —
    /// the only clue a file gives that its hidden rows were hidden by filtering.
    /// </summary>
    void OnAutoFilter(RangeRef range, bool hasCriteria);

    /// <summary>One data validation rule.</summary>
    void OnDataValidation(DataValidation validation);

    /// <summary>One conditional formatting rule.</summary>
    void OnConditionalFormat(ConditionalFormatRule rule);

    /// <summary>The sheet's print settings, to be filled in from a captured element.</summary>
    PageSetup PageSetup { get; }

    /// <summary>The sheet's protection, to be filled in from a captured element.</summary>
    SheetProtection Protection { get; }

    /// <summary>
    /// A hyperlink over a range. External targets arrive as a relationship id to resolve
    /// against the sheet's relationships; internal ones carry a <paramref name="location"/>.
    /// </summary>
    void OnHyperlink(RangeRef range, string? relationshipId, string? location, string? tooltip, string? display);

    /// <summary>
    /// A shared or array formula master. <paramref name="sharedIndex"/> is the <c>si</c>
    /// attribute for shared groups, letting members be matched to this master.
    /// </summary>
    void OnFormulaGroup(FormulaGroup group, int? sharedIndex);

    /// <summary>A cell that repeats a shared formula identified by <paramref name="sharedIndex"/>.</summary>
    void OnSharedFormulaMember(CellRef cell, int sharedIndex);

    /// <summary>Attributes of a cell's formula element the model does not represent.</summary>
    void OnFormulaMarkup(CellRef cell, FormulaMarkup markup);

    /// <summary>
    /// Attributes of a cell the model does not represent, chiefly the <c>cm</c> and <c>vm</c>
    /// indexes tying it to the metadata part.
    /// </summary>
    void OnCellMarkup(CellRef cell, List<KeyValuePair<string, string>> attributes);

    /// <summary>Everything in the part the model does not represent, to be replayed on save.</summary>
    void OnMarkup(WorksheetMarkup markup);
}
