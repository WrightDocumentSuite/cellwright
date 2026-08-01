using Cellwright.Xls.Cfb;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Carries the title, author and the rest between the property set streams of a compound file
/// and the workbook's own properties.
/// </summary>
/// <remarks>
/// The legacy format keeps these outside the BIFF stream entirely, in two streams beside it
/// that hold OLE property sets. Nothing in the workbook stream points at them, which is why a
/// reader that only follows records never finds them and quietly loses everything a user typed
/// into the document's properties dialog.
/// </remarks>
internal static class DocumentProperties
{
    private static readonly Guid SummaryInformation = new("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
    private static readonly Guid DocumentSummaryInformation = new("D5CDD502-2E9C-101B-9397-08002B2CF9AE");
    private static readonly Guid UserDefinedProperties = new("D5CDD505-2E9C-101B-9397-08002B2CF9AE");

    private const string SummaryStream = "\u0005SummaryInformation";
    private const string DocumentSummaryStream = "\u0005DocumentSummaryInformation";

    // SummaryInformation identifiers (MS-OLEPS 2.25.2).
    private const uint Title = 0x02;
    private const uint Subject = 0x03;
    private const uint Author = 0x04;
    private const uint Keywords = 0x05;
    private const uint Comments = 0x06;
    private const uint LastAuthor = 0x08;
    private const uint Created = 0x0C;
    private const uint LastSaved = 0x0D;
    private const uint Application = 0x12;

    // DocumentSummaryInformation identifiers (MS-OSHARED 2.3.3.2.2.1).
    private const uint Category = 0x02;
    private const uint Manager = 0x0E;
    private const uint Company = 0x0F;
    private const uint ContentStatus = 0x1B;

    /// <summary>Reads both property streams of a container into the workbook.</summary>
    /// <remarks>
    /// The streams are looked up in the root storage and nowhere else. An embedded Word or
    /// PowerPoint document is a storage of its own inside the same container, and it carries
    /// property streams under exactly these names — so a search by name across the whole
    /// directory hands back the title and author of the attachment instead of the workbook's.
    /// </remarks>
    public static void Read(CompoundFile container, WorkbookProperties properties)
    {
        if (container.TryReadPath(SummaryStream) is { } summary)
            ReadSummary(PropertySetStream.Read(summary), properties);

        if (container.TryReadPath(DocumentSummaryStream) is { } document)
            ReadDocumentSummary(PropertySetStream.Read(document), properties);
    }

    private static void ReadSummary(
        Dictionary<Guid, PropertySetStream.Section> sets, WorkbookProperties properties)
    {
        if (!sets.TryGetValue(SummaryInformation, out PropertySetStream.Section section))
            return;

        properties.Title = Text(section, Title) ?? properties.Title;
        properties.Subject = Text(section, Subject) ?? properties.Subject;
        properties.Author = Text(section, Author) ?? properties.Author;
        properties.Keywords = Text(section, Keywords) ?? properties.Keywords;
        properties.Description = Text(section, Comments) ?? properties.Description;
        properties.LastModifiedBy = Text(section, LastAuthor) ?? properties.LastModifiedBy;
        properties.Application = Text(section, Application) ?? properties.Application;
        properties.Created = Moment(section, Created) ?? properties.Created;
        properties.Modified = Moment(section, LastSaved) ?? properties.Modified;
    }

    private static void ReadDocumentSummary(
        Dictionary<Guid, PropertySetStream.Section> sets, WorkbookProperties properties)
    {
        if (sets.TryGetValue(DocumentSummaryInformation, out PropertySetStream.Section section))
        {
            properties.Category = Text(section, Category) ?? properties.Category;
            properties.Manager = Text(section, Manager) ?? properties.Manager;
            properties.Company = Text(section, Company) ?? properties.Company;
            properties.ContentStatus = Text(section, ContentStatus) ?? properties.ContentStatus;
        }

        // The user-defined set shares this stream, and its properties are named by a dictionary
        // rather than by fixed identifiers.
        if (!sets.TryGetValue(UserDefinedProperties, out PropertySetStream.Section custom))
            return;

        foreach ((uint id, string name) in custom.Names)
        {
            if (custom.Values.TryGetValue(id, out object? value) && Convert(value) is { } converted)
                properties.Custom[name] = converted;
        }
    }

    private static string? Text(in PropertySetStream.Section section, uint id) =>
        section.Values.TryGetValue(id, out object? value) && value is string { Length: > 0 } text ? text : null;

    private static DateTime? Moment(in PropertySetStream.Section section, uint id) =>
        section.Values.TryGetValue(id, out object? value) && value is DateTime moment ? moment : null;

    /// <summary>Turns a property value into the cell value a custom property is modelled as.</summary>
    private static CellValue? Convert(object value) => value switch
    {
        string text => CellValue.Text(text),
        int number => CellValue.Number(number),
        double number => CellValue.Number(number),
        bool flag => CellValue.Boolean(flag),
        DateTime moment => CellValue.DateTime(moment),
        _ => (CellValue?)null,
    };
}
