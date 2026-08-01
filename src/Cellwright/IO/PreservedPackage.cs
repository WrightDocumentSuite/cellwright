using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Cellwright.IO;

/// <summary>
/// Everything a loaded package holds that the document model does not represent: raw part
/// bytes, the relationship graph and the content-type map. Saving copies these through
/// untouched, so editing one cell of a real workbook cannot cost it its charts, pivot
/// tables, images, printer settings or macros.
/// </summary>
internal sealed class PreservedPackage
{
    /// <summary>Content types of the source package, merged into the ones we generate.</summary>
    public required ContentTypeMap ContentTypes { get; init; }

    /// <summary>Where the workbook part lived. Reused on save so relative targets keep resolving.</summary>
    public required string WorkbookPartPath { get; init; }

    /// <summary>Content type of the workbook part; macro-enabled workbooks differ from plain ones.</summary>
    public required string WorkbookContentType { get; init; }

    /// <summary>Original location of the styles part, or <see langword="null"/> when the package had none.</summary>
    public string? StylesPartPath { get; init; }

    /// <summary>Original location of the shared-string part, or <see langword="null"/>.</summary>
    public string? SharedStringsPartPath { get; init; }

    /// <summary>
    /// Where the document properties lived. The canonical <c>docProps</c> names are only a
    /// convention: a package written through System.IO.Packaging keeps its core properties
    /// under <c>package/services</c> with a generated name, and moving them would look to
    /// anything comparing the two files like the part had been dropped.
    /// </summary>
    public string? CorePropertiesPartPath { get; init; }

    /// <inheritdoc cref="CorePropertiesPartPath"/>
    public string? ExtendedPropertiesPartPath { get; init; }

    /// <inheritdoc cref="CorePropertiesPartPath"/>
    public string? CustomPropertiesPartPath { get; init; }

    /// <summary>Uninterpreted markup of <c>workbook.xml</c>.</summary>
    public WorkbookMarkup Markup { get; } = new();

    /// <summary>Uninterpreted sections of <c>styles.xml</c>, such as the differential formats tables index into.</summary>
    public Styles.StylesMarkup StyleMarkup { get; init; } = new();

    /// <summary>Chart, dialog and macro sheets: listed in the workbook but outside the grid model.</summary>
    public IReadOnlyList<ForeignSheet> ForeignSheets { get; init; } = [];

    /// <summary>Raw bytes of every part the model does not regenerate, keyed by part name.</summary>
    public Dictionary<string, byte[]> Parts { get; } = new(OpcNames.Instance);

    /// <summary>Relationships of every part that had any, keyed by the source part name.</summary>
    public Dictionary<string, IReadOnlyList<OpcRelationship>> Relationships { get; } =
        new(OpcNames.Instance);

    /// <summary>Relationships declared by a part; empty when it declared none.</summary>
    public IReadOnlyList<OpcRelationship> RelationshipsOf(string partPath) =>
        Relationships.TryGetValue(partPath, out IReadOnlyList<OpcRelationship>? relationships) ? relationships : [];

    /// <summary>
    /// Parts nothing referenced in the source package either. Garbage collection only removes
    /// what a deleted sheet used to own; a part that arrived unanchored is copied as it was,
    /// because dropping it would be a change nobody asked for.
    /// </summary>
    public IReadOnlySet<string> OrphanParts { get; set; } = new HashSet<string>(OpcNames.Instance);

    /// <summary>
    /// Part names reachable by following relationships from the given roots. Saving copies
    /// only these, so parts that only a deleted sheet referenced are collected rather than
    /// carried into the new file as orphans.
    /// </summary>
    public HashSet<string> Reachable(IEnumerable<string> roots)
    {
        var visited = new HashSet<string>(OpcNames.Instance);
        var pending = new Queue<string>(roots);
        while (pending.TryDequeue(out string? part))
        {
            foreach (OpcRelationship relationship in RelationshipsOf(part))
            {
                if (relationship.IsExternal)
                    continue;

                string target = OpcPath.Resolve(part, relationship.Target);
                if (visited.Add(target))
                    pending.Enqueue(target);
            }
        }

        return visited;
    }

    /// <summary>
    /// Points the parts that name a worksheet in their own markup at its new name: the
    /// <c>c:f</c> series references of charts and the <c>worksheetSource</c> of pivot caches.
    /// Both spell the sheet out in text, so a rename would otherwise leave a chart plotting
    /// nothing.
    /// </summary>
    public void RetargetSheetName(string oldName, string newName)
    {
        foreach (string partPath in Parts.Keys)
        {
            bool isChart = partPath.Contains("/charts/", StringComparison.OrdinalIgnoreCase);
            bool isPivot = partPath.Contains("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase);
            if (!isChart && !isPivot)
                continue;

            byte[] content = Parts[partPath];
            byte[] rewritten = RetargetXml(content, oldName, newName, isChart);
            if (!ReferenceEquals(content, rewritten))
                Parts[partPath] = rewritten;
        }
    }

    /// <summary>
    /// Rewrites chart formulas or pivot sources by expanded XML name. Prefixes, quote styles and
    /// whitespace are producer choices and cannot be used to identify an OOXML element.
    /// </summary>
    private static byte[] RetargetXml(byte[] content, string oldName, string newName, bool chart)
    {
        XDocument document;
        try
        {
            using var input = new MemoryStream(content, writable: false);
            document = XDocument.Load(input, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            // Preservation must never turn an already malformed foreign part into another file.
            return content;
        }

        bool changed = false;
        if (chart)
        {
            foreach (XElement formula in document.Descendants().Where(static element =>
                         element.Name.LocalName == "f" && IsChartNamespace(element.Name.NamespaceName)))
            {
                string rewritten = FormulaText.RenameSheet(formula.Value, oldName, newName)!;
                if (string.Equals(formula.Value, rewritten, StringComparison.Ordinal))
                    continue;
                formula.Value = rewritten;
                changed = true;
            }
        }
        else
        {
            foreach (XElement source in document.Descendants().Where(static element =>
                         element.Name.LocalName == "worksheetSource" &&
                         XlsxSchema.IsSpreadsheetNamespace(element.Name.NamespaceName)))
            {
                XAttribute? sheet = source.Attribute("sheet");
                if (sheet is null || !string.Equals(sheet.Value, oldName, StringComparison.Ordinal))
                    continue;
                sheet.Value = newName;
                changed = true;
            }
        }

        if (!changed)
            return content;

        using var output = new MemoryStream(content.Length);
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration is null,
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings))
            document.Save(writer);
        return output.ToArray();

        static bool IsChartNamespace(string value) =>
            value is "http://schemas.openxmlformats.org/drawingml/2006/chart" or
                "http://purl.oclc.org/ooxml/drawingml/chart";
    }
}
