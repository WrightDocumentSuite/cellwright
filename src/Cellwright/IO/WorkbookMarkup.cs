namespace Cellwright.IO;

/// <summary>
/// The parts of <c>workbook.xml</c> that the model does not represent, split into the slots
/// the schema defines rather than kept as one blob: the elements we do generate
/// (<c>workbookPr</c>, <c>sheets</c>, <c>definedNames</c>, <c>calcPr</c>) sit between them.
/// </summary>
internal sealed class WorkbookMarkup
{
    /// <summary>Extra attributes of the root element, e.g. <c>mc:Ignorable</c> and namespace declarations.</summary>
    public List<KeyValuePair<string, string>>? RootAttributes { get; set; }

    /// <summary>Attributes of <c>workbookPr</c> other than <c>date1904</c>, which the model owns.</summary>
    public List<KeyValuePair<string, string>>? WorkbookProperties { get; set; }

    /// <summary>Attributes of <c>calcPr</c> other than <c>fullCalcOnLoad</c>, which the model owns.</summary>
    public List<KeyValuePair<string, string>>? CalculationProperties { get; set; }

    /// <summary><c>fileVersion</c>, <c>fileSharing</c> — everything preceding <c>workbookPr</c>.</summary>
    public string? BeforeProperties { get; set; }

    /// <summary><c>workbookProtection</c>, <c>bookViews</c> — between <c>workbookPr</c> and <c>sheets</c>.</summary>
    public string? AfterProperties { get; set; }

    /// <summary><c>functionGroups</c>, <c>externalReferences</c> — between <c>sheets</c> and <c>definedNames</c>.</summary>
    public string? BeforeDefinedNames { get; set; }

    /// <summary><c>oleSize</c>, <c>customWorkbookViews</c>, <c>pivotCaches</c>, <c>extLst</c> and friends.</summary>
    public string? AfterCalculationProperties { get; set; }
}
