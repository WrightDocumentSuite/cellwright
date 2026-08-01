namespace Cellwright.IO;

/// <summary>
/// Well-known content types, relationship types and XML namespaces of the xlsx package graph.
/// </summary>
internal static class XlsxSchema
{
    // Content types.
    public const string ContentTypeWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    public const string ContentTypeWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    public const string ContentTypeSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    public const string ContentTypeStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    public const string ContentTypeRelationships = "application/vnd.openxmlformats-package.relationships+xml";
    public const string ContentTypeCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
    public const string ContentTypeExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    public const string ContentTypeCustomProperties = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    public const string ContentTypeComments = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    public const string ContentTypeTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    public const string ContentTypeVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    public const string ContentTypeDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";

    // Relationship types.
    public const string RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    public const string RelWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    public const string RelSharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    public const string RelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    public const string RelCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    public const string RelExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    public const string RelCustomProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    public const string RelCalcChain = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    public const string RelHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    public const string RelComments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    /// <summary>
    /// The threaded comments of a sheet (MS-XLSX 2.1.17). An extension rather than part of
    /// ISO/IEC 29500, so the namespace names Microsoft and a year rather than the standard.
    /// </summary>
    public const string RelThreadedComments =
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";

    /// <summary>The people who wrote a workbook's threaded comments (MS-XLSX 2.1.18).</summary>
    public const string RelPersons = "http://schemas.microsoft.com/office/2017/10/relationships/person";

    /// <inheritdoc cref="RelThreadedComments"/>
    public const string ContentTypeThreadedComments = "application/vnd.ms-excel.threadedcomments+xml";

    /// <inheritdoc cref="RelPersons"/>
    public const string ContentTypePersons = "application/vnd.ms-excel.person+xml";

    /// <summary>The cell metadata part, which is what tells a dynamic array from a CSE formula.</summary>
    public const string RelSheetMetadata =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";

    /// <inheritdoc cref="RelSheetMetadata"/>
    public const string ContentTypeSheetMetadata =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";

    /// <summary>Canonical metadata part path, of which a workbook has at most one.</summary>
    public const string PartMetadata = "/xl/metadata.xml";
    public const string RelVmlDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    public const string RelTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    public const string RelDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

    /// <summary>The chart part a drawing's graphic frame points at.</summary>
    public const string RelChart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

    /// <inheritdoc cref="RelChart"/>
    public const string ContentTypeChart =
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

    /// <summary>Canonical chart part path for a 1-based chart index.</summary>
    public static string ChartPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/charts/chart{index}.xml");
    public const string RelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    public const string RelTheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    // XML namespaces.
    public const string NsSpreadsheetMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string NsStrictSpreadsheetMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    public const string NsSpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public const string NsDrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string NsDocumentRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string NsPackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>The relationships namespace of a Strict file, where <c>r:id</c> also lives.</summary>
    public const string NsStrictDocumentRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    /// <summary>Markup compatibility (ISO/IEC 29500-3), home of <c>Ignorable</c> and <c>MustUnderstand</c>.</summary>
    public const string NsMarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>Whether a namespace is the transitional or Strict SpreadsheetML vocabulary.</summary>
    public static bool IsSpreadsheetNamespace(string value) =>
        value is NsSpreadsheetMain or NsStrictSpreadsheetMain;

    private const string StrictPackageRelationships = "http://purl.oclc.org/ooxml/package/relationships";

    /// <summary>
    /// How LibreOffice names the core-properties relationship: under the officeDocument
    /// authority rather than the package one, and with a small d. It means the same thing, and
    /// enough files in the wild carry it that not recognizing it is a bug of ours, not theirs.
    /// </summary>
    private const string LibreOfficeCoreProperties =
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";

    /// <summary>
    /// Rewrites a Strict relationship type as the transitional one it means. The two
    /// conformance classes of ISO/IEC 29500 name the same roles under different authorities, so
    /// recognizing both is all it takes to read a Strict file.
    /// </summary>
    public static string Canonical(string type)
    {
        if (string.Equals(type, LibreOfficeCoreProperties, StringComparison.Ordinal))
            return RelCoreProperties;
        if (Rebase(type, NsStrictDocumentRelationships, NsDocumentRelationships) is { } document)
            return document;
        return Rebase(type, StrictPackageRelationships, NsPackageRelationships) ?? type;
    }

    /// <summary>Whether a part name is one of the ones under <c>docProps</c> that the writer generates.</summary>
    public static bool IsDocumentPropertiesPart(string partPath) =>
        partPath.Equals(PartCoreProperties, StringComparison.OrdinalIgnoreCase) ||
        partPath.Equals(PartExtendedProperties, StringComparison.OrdinalIgnoreCase) ||
        partPath.Equals(PartCustomProperties, StringComparison.OrdinalIgnoreCase);

    private static string? Rebase(string type, string strictPrefix, string transitionalPrefix) =>
        type.StartsWith(strictPrefix, StringComparison.Ordinal)
            ? transitionalPrefix + type[strictPrefix.Length..]
            : null;

    // Canonical part paths used when writing new packages.
    public const string PartWorkbook = "/xl/workbook.xml";
    public const string PartStyles = "/xl/styles.xml";
    public const string PartSharedStrings = "/xl/sharedStrings.xml";
    public const string PartCoreProperties = "/docProps/core.xml";
    public const string PartExtendedProperties = "/docProps/app.xml";
    public const string PartCustomProperties = "/docProps/custom.xml";

    /// <summary>Canonical worksheet part path for a 1-based sheet index.</summary>
    public static string WorksheetPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/worksheets/sheet{index}.xml");

    /// <summary>Canonical comments part path for a 1-based sheet index.</summary>
    public static string CommentsPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/comments{index}.xml");

    /// <summary>Canonical threaded comments part path for a 1-based sheet index.</summary>
    public static string ThreadedCommentsPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/threadedComments/threadedComment{index}.xml");

    /// <summary>Canonical persons part path, of which a workbook has at most one.</summary>
    public const string PartPersons = "/xl/persons/person.xml";

    /// <summary>Canonical legacy VML drawing part path for a 1-based sheet index.</summary>
    public static string VmlDrawingPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/drawings/vmlDrawing{index}.vml");

    /// <summary>Canonical table part path for a 1-based table index.</summary>
    public static string TablePart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/tables/table{index}.xml");

    /// <summary>Canonical drawing part path for a 1-based sheet index.</summary>
    public static string DrawingPart(int index) =>
        string.Create(null, stackalloc char[64], $"/xl/drawings/drawing{index}.xml");

    /// <summary>Canonical media part path for a 1-based image index and an extension without the dot.</summary>
    public static string MediaPart(int index, string extension) =>
        string.Create(null, stackalloc char[64], $"/xl/media/image{index}.{extension}");
}
