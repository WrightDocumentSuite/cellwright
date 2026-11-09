namespace Cellwright.Sanitization;

/// <summary>Kinds of content a sanitized copy can leave out.</summary>
[Flags]
public enum ContentCategories
{
    /// <summary>Nothing.</summary>
    None = 0,

    /// <summary>VBA projects and Excel 4.0 macro sheets.</summary>
    Macros = 1,

    /// <summary>OLE objects, embedded packages and ActiveX controls.</summary>
    EmbeddedObjects = 2,

    /// <summary>
    /// Links to other workbooks and every relationship that points outside the package, hyperlinks
    /// included. A formula that reads another workbook keeps its cached value and loses the formula.
    /// </summary>
    ExternalLinks = 4,

    /// <summary>Custom XML data parts.</summary>
    CustomXml = 8,

    /// <summary>Digital signatures and their certificates.</summary>
    Signatures = 16,

    /// <summary>All of the above.</summary>
    All = Macros | EmbeddedObjects | ExternalLinks | CustomXml | Signatures,
}

/// <summary>What a <see cref="SanitizedItem"/> describes.</summary>
public enum SanitizedItemKind
{
    /// <summary>A package part.</summary>
    Part,

    /// <summary>A relationship from one part to a part or an address.</summary>
    Relationship,

    /// <summary>Markup inside a part.</summary>
    Markup,
}

/// <summary>One thing a sanitized copy left out, or kept.</summary>
/// <param name="Category">The category it belongs to.</param>
/// <param name="Kind">What it is.</param>
/// <param name="Location">The part or address concerned.</param>
/// <param name="Detail">The relationship type, element name or field instruction.</param>
public readonly record struct SanitizedItem(
    ContentCategories Category, SanitizedItemKind Kind, string Location, string Detail);

/// <summary>What a sanitized save did, so that the result can be checked rather than assumed.</summary>
public sealed class SanitizeReport
{
    internal SanitizeReport(IReadOnlyList<SanitizedItem> removed, IReadOnlyList<SanitizedItem> retained)
    {
        Removed = removed;
        Retained = retained;
    }

    /// <summary>Everything left out of the copy.</summary>
    public IReadOnlyList<SanitizedItem> Removed { get; }

    /// <summary>
    /// Content in a known category that is still in the copy because its category was not asked for.
    /// </summary>
    public IReadOnlyList<SanitizedItem> Retained { get; }
}
