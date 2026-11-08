using System.Xml;
using System.Xml.Linq;
using Cellwright.Xml;

namespace Cellwright.IO;

/// <summary>Rich shared-string markup and the plain text it originally represented.</summary>
internal readonly record struct RichStringMarkup(string Text, string Markup);

/// <summary>
/// Read-only semantic view of the extension-only conditional formats a worksheet carries.
/// The raw <c>extLst</c> remains the source of truth for saving; these rules exist only so
/// evaluators can apply the subset the model understands without exposing it as editable state.
/// </summary>
internal sealed class ExtensionConditionalFormatProjection(
    IReadOnlyList<ConditionalFormatRule> rules,
    bool hasUnsupportedRules)
{
    /// <summary>Rules that can be evaluated without changing their extension representation.</summary>
    public IReadOnlyList<ConditionalFormatRule> Rules { get; } = [.. rules];

    /// <summary>Whether at least one extension rule could not be represented by this view.</summary>
    public bool HasUnsupportedRules { get; } = hasUnsupportedRules;
}

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

    /// <summary>
    /// Priorities owned by extension-only x14 conditional-format rules. Base rules regenerated
    /// from the model must not reuse them because priority is worksheet-wide across both forms.
    /// </summary>
    public IReadOnlySet<int>? ExtensionConditionalFormatPriorities { get; set; }

    /// <summary>Whether the trailing extension list contains x14 conditional formatting.</summary>
    public bool HasExtensionConditionalFormats { get; set; }

    /// <summary>
    /// Semantic view of supported extension-only rules. It is deliberately separate from
    /// <see cref="ConditionalFormats"/> so regenerating editable base rules cannot duplicate an
    /// extension rule or write it in the wrong namespace.
    /// </summary>
    public ExtensionConditionalFormatProjection? ExtensionConditionalFormats { get; set; }

    /// <summary>
    /// Whether some extension conditional formatting remains unavailable to the evaluator.
    /// A missing projection while x14 content is still present means none of it can be applied.
    /// </summary>
    public bool HasUnsupportedExtensionConditionalFormats =>
        HasExtensionConditionalFormats &&
        (ExtensionConditionalFormats is null || ExtensionConditionalFormats.HasUnsupportedRules);

    /// <summary>
    /// Whether the worksheet carried source-only markup that is not rebuilt completely from the
    /// copy model. The ordinary drawing pointer is excluded because pictures and charts receive
    /// a fresh drawing graph; unmodelled anchors are tracked separately on the worksheet.
    /// </summary>
    public bool HasCopySensitiveContent =>
        RootAttributes is not null ||
        SheetProperties is not null || SheetPropertyChildren is not null ||
        SheetViewAttributes is not null || SheetViewChildren is not null ||
        SheetFormatProperties is not null || BeforeAutoFilter is not null ||
        AfterProtection is not null || AutoFilter is not null || BeforeMerges is not null ||
        AfterMerges is not null || ConditionalFormats is not null || DataValidations is not null ||
        PageSetup is not null || PageSetupAttributes is not null || PageSetupProperties is not null ||
        PrintSetup is not null || LegacyDrawing is not null ||
        AfterLegacyDrawing is not null || TableParts is not null || Extensions is not null;

    /// <summary>
    /// Stops using a semantic view that could not be trusted or safely removed. Structural edits
    /// use <see cref="RemoveExtensionConditionalFormats"/> so stale raw ranges are not saved.
    /// </summary>
    public void InvalidateExtensionConditionalFormats() => ExtensionConditionalFormats = null;

    /// <summary>
    /// Removes the x14 conditional-format extension after an edit that moved its unmodelled
    /// ranges or formulas. Other extension elements and the outer list remain intact.
    /// </summary>
    /// <returns>Whether an x14 conditional-format container was removed.</returns>
    public bool RemoveExtensionConditionalFormats()
    {
        if (!HasExtensionConditionalFormats || string.IsNullOrWhiteSpace(Extensions))
        {
            InvalidateExtensionConditionalFormats();
            return false;
        }

        const string x14Namespace =
            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        try
        {
            XmlReaderSettings settings = XmlDefaults.ReaderSettings.Clone();
            settings.IgnoreWhitespace = false;
            using var xml = XmlReader.Create(new StringReader(Extensions), settings);
            XDocument document = XDocument.Load(xml, System.Xml.Linq.LoadOptions.PreserveWhitespace);
            XName name = XName.Get("conditionalFormattings", x14Namespace);
            XElement[] containers = [.. document.Descendants(name)];
            if (containers.Length == 0)
            {
                InvalidateExtensionConditionalFormats();
                return false;
            }

            XElement[] parents = [.. containers
                .Select(static container => container.Parent)
                .OfType<XElement>()
                .Distinct()];
            foreach (XElement container in containers)
                container.Remove();

            // The dedicated Office extension normally has no other child. Remove that now-empty
            // wrapper, but retain it when a producer placed unrelated extension data beside CF.
            foreach (XElement parent in parents)
            {
                if (!parent.Elements().Any())
                    parent.Remove();
            }

            Extensions = document.Root is { } root && root.Elements().Any()
                ? document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                : null;
            HasExtensionConditionalFormats = false;
            ExtensionConditionalFormatPriorities = null;
            ExtensionConditionalFormats = null;
            return true;
        }
        catch (XmlException)
        {
            // Captured worksheet markup is well formed in the ordinary loader path. If a caller
            // supplies malformed internal markup, keeping it opaque is safer than deleting other
            // extensions while guessing at element boundaries.
            InvalidateExtensionConditionalFormats();
            return false;
        }
    }
}
