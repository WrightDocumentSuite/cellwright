namespace Cellwright.IO;

/// <summary>Rich shared-string markup and the plain text it originally represented.</summary>
internal readonly record struct RichStringMarkup(string Text, string Markup);

/// <summary>
/// The parts of a worksheet the model does not represent, kept in the slots the schema
/// defines. A worksheet is always regenerated (its cells are the whole point), so unlike
/// other parts it cannot simply be copied — sheet protection, conditional formatting, print
/// setup, drawings and OLE objects have to be replayed around the content we do own.
/// </summary>
internal sealed class WorksheetMarkup
{
    /// <summary>Extra attributes of the root element, e.g. <c>mc:Ignorable</c> and namespace declarations.</summary>
    public List<KeyValuePair<string, string>>? RootAttributes { get; set; }

    /// <summary>Attributes of <c>sheetPr</c>; the model owns none of them.</summary>
    public List<KeyValuePair<string, string>>? SheetProperties { get; set; }

    /// <summary>Children of <c>sheetPr</c> other than <c>tabColor</c> and <c>pageSetUpPr</c>, e.g. <c>outlinePr</c>.</summary>
    public string? SheetPropertyChildren { get; set; }

    /// <summary>Attributes of the first <c>sheetView</c> other than the ones the model owns.</summary>
    public List<KeyValuePair<string, string>>? SheetViewAttributes { get; set; }

    /// <summary>Children of the first <c>sheetView</c> other than <c>pane</c>: selections and extensions.</summary>
    public string? SheetViewChildren { get; set; }

    /// <summary>The <c>sheetFormatPr</c> element: default row height, outline levels.</summary>
    public string? SheetFormatProperties { get; set; }

    /// <summary><c>sheetCalcPr</c> — what sits between the cells and the protection.</summary>
    public string? BeforeAutoFilter { get; set; }

    /// <summary><c>scenarios</c> — what sits between the protection and the auto filter.</summary>
    public string? AfterProtection { get; set; }

    /// <summary>The <c>autoFilter</c> element, replayed while the model has not changed its range.</summary>
    public string? AutoFilter { get; set; }

    /// <summary>Range the captured <c>autoFilter</c> covered, used to tell whether it still applies.</summary>
    public RangeRef? AutoFilterRange { get; set; }

    /// <summary><c>sortState</c>, <c>dataConsolidate</c>, <c>customSheetViews</c>.</summary>
    public string? BeforeMerges { get; set; }

    /// <summary><c>phoneticPr</c> and anything else between the merges and the rules.</summary>
    public string? AfterMerges { get; set; }

    /// <summary>The <c>conditionalFormatting</c> blocks, replayed while the model has not changed them.</summary>
    public string? ConditionalFormats { get; set; }

    /// <summary>The <c>dataValidations</c> element, replayed while the model has not changed it.</summary>
    public string? DataValidations { get; set; }

    /// <summary>
    /// <c>printOptions</c> through <c>colBreaks</c>, replayed while the model has not changed
    /// the print settings.
    /// </summary>
    public string? PageSetup { get; set; }

    /// <summary>Attributes of <c>pageSetup</c> the model does not represent, chiefly the printer settings link.</summary>
    public List<KeyValuePair<string, string>>? PageSetupAttributes { get; set; }

    /// <summary>The <c>pageSetUpPr</c> child of <c>sheetPr</c>, kept while the settings are untouched.</summary>
    public string? PageSetupProperties { get; set; }

    /// <summary><c>customProperties</c> through <c>smartTags</c>: cell watches, ignored errors.</summary>
    public string? PrintSetup { get; set; }

    /// <summary>The <c>drawing</c> element pointing at the sheet's shapes and charts.</summary>
    public string? Drawing { get; set; }

    /// <summary>
    /// The <c>legacyDrawing</c> element. Kept apart from the rest because it points at the VML
    /// that gives notes their boxes, and editing notes means replacing it.
    /// </summary>
    public string? LegacyDrawing { get; set; }

    /// <summary><c>legacyDrawingHF</c>, <c>drawingHF</c>, <c>picture</c>, <c>oleObjects</c>, <c>controls</c>, <c>webPublishItems</c>.</summary>
    public string? AfterLegacyDrawing { get; set; }

    /// <summary>The <c>tableParts</c> element, replayed while the model has not changed it.</summary>
    public string? TableParts { get; set; }

    /// <summary>The trailing <c>extLst</c> element.</summary>
    public string? Extensions { get; set; }
}
