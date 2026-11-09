using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Cellwright.Sanitization;

/// <summary>The parts of sanitizing that depend on SpreadsheetML.</summary>
internal sealed partial class XlsxContentSanitizer(ContentCategories remove) : OpcContentSanitizer(remove)
{
    private static readonly HashSet<string> Prunable = new(StringComparer.Ordinal)
    {
        "oleObjects", "controls", "hyperlinks", "externalReferences",
        "definedNames", "dataValidations", "conditionalFormatting",
    };

    private bool _macroSheetRemoved;

    protected override IReadOnlySet<string> PruneWhenEmpty => Prunable;

    protected override string? PlainContentType(string contentType) => contentType switch
    {
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml" or "application/vnd.ms-excel.addin.macroEnabled.main+xml" =>
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml" =>
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        _ => null,
    };

    protected override ContentCategories Classify(Relationship relationship) => relationship.TypeName switch
    {
        "xlMacrosheet" or "xlIntlMacrosheet" when !relationship.IsExternal => ContentCategories.Macros,
        "externalLink" => ContentCategories.ExternalLinks,
        _ => base.Classify(relationship),
    };

    protected override void RemoveReference(XElement element, ReferenceScope scope)
    {
        if (element.Name.LocalName == "sheet")
        {
            ShiftSheetIndexes(element);
            _macroSheetRemoved = true;
        }

        scope.Remove(element);
    }

    /// <summary>Names and views refer to sheets by position, which a removed sheet changes.</summary>
    private static void ShiftSheetIndexes(XElement sheet)
    {
        int removed = sheet.ElementsBeforeSelf().Count();

        foreach (XElement name in sheet.Document!.Descendants().Where(static e => e.Name.LocalName == "definedName").ToArray())
        {
            if (name.Attribute("localSheetId") is not { } attribute || !int.TryParse(attribute.Value, out int index))
                continue;

            if (index == removed)
            {
                XElement? names = name.Parent;
                name.Remove();
                if (names is { HasElements: false })
                    names.Remove();
            }
            else if (index > removed)
            {
                attribute.Value = (index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        foreach (XElement view in sheet.Document.Descendants().Where(static e => e.Name.LocalName == "workbookView"))
        {
            ShiftIndex(view.Attribute("activeTab"), removed);
            ShiftIndex(view.Attribute("firstSheet"), removed);
        }
    }

    private static void ShiftIndex(XAttribute? attribute, int removed)
    {
        if (attribute is null || !int.TryParse(attribute.Value, out int index) || index < removed)
            return;

        attribute.Value = Math.Max(0, index == removed ? 0 : index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    protected override void ApplyFormatRules()
    {
        bool formulasChanged = false;
        if ((Selected & ContentCategories.ExternalLinks) != 0)
        {
            foreach ((string _, string sheet) in TargetsOfType("worksheet"))
                formulasChanged |= UnlinkSheet(sheet);

            foreach (string workbook in TargetsOfType("worksheet").Select(static pair => pair.Source).Distinct(StringComparer.OrdinalIgnoreCase))
                RemoveLinkedNames(workbook);
        }

        if (formulasChanged || _macroSheetRemoved)
            DropCalculationChain(formulasChanged ? ContentCategories.ExternalLinks : ContentCategories.Macros);
    }

    /// <summary>
    /// A formula that reads another workbook cannot outlive the link it reads through: Excel
    /// refuses an index into external references that are not there. Its last value stays.
    /// </summary>
    private bool UnlinkSheet(string sheet)
    {
        if (!PartPaths.Contains(sheet, StringComparer.OrdinalIgnoreCase) || !ReadText(sheet).Contains('['))
            return false;

        XDocument document = ReadXml(sheet);
        var scope = new ReferenceScope(this, sheet, PruneWhenEmpty) { Category = ContentCategories.ExternalLinks };

        XElement[] formulas = [.. document.Descendants().Where(static e => e.Name.LocalName == "f" && ExternalIndex().IsMatch(e.Value))];
        foreach (XElement formula in formulas)
            formula.Remove();

        if (formulas.Length > 0)
            ReportRemoved(ContentCategories.ExternalLinks, SanitizedItemKind.Markup, sheet, $"{formulas.Length} formula(s) replaced by their cached values");

        XElement[] rules = [.. document.Descendants().Where(static e => e.Name.LocalName is "formula" or "formula1" or "formula2" && ExternalIndex().IsMatch(e.Value))];
        foreach (XElement rule in rules)
        {
            if (rule.Parent is { Document: not null } owner && owner.Name.LocalName is "cfRule" or "dataValidation")
                scope.Remove(owner);
        }

        foreach (XElement validations in document.Descendants().Where(static e => e.Name.LocalName == "dataValidations"))
            validations.Attribute("count")?.SetValue(validations.Elements().Count());

        if (formulas.Length == 0 && rules.Length == 0)
            return false;

        WriteXml(sheet, document);
        return formulas.Length > 0;
    }

    private void RemoveLinkedNames(string workbook)
    {
        XDocument document = ReadXml(workbook);
        var scope = new ReferenceScope(this, workbook, PruneWhenEmpty) { Category = ContentCategories.ExternalLinks };

        XElement[] names = [.. document.Descendants().Where(static e => e.Name.LocalName == "definedName" && ExternalIndex().IsMatch(e.Value))];
        foreach (XElement name in names)
            scope.Remove(name);

        if (names.Length > 0)
            WriteXml(workbook, document);
    }

    /// <summary>The calculation chain lists formula cells; Excel rebuilds it when it is absent.</summary>
    private void DropCalculationChain(ContentCategories cause)
    {
        foreach ((string source, string _) in TargetsOfType("calcChain"))
        {
            foreach (Relationship relationship in RelationshipsOf(source).Where(static r => r.TypeName == "calcChain").ToArray())
                DropRelationship(relationship, cause);
        }
    }

    /// <summary>
    /// <c>[1]</c> indexes an external reference. A structured reference to a column named with
    /// digits, as in <c>Sales[2024]</c>, has a table name right before the bracket.
    /// </summary>
    [GeneratedRegex(@"(?<![\p{L}\p{N}_.\]])\[\d+\]", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex ExternalIndex();
}
