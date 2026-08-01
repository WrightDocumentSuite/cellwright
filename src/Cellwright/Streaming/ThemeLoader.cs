using Cellwright.IO;

namespace Cellwright.Streaming;

/// <summary>
/// Finds and reads the theme part a workbook points at.
/// </summary>
/// <remarks>
/// Shared by both package formats for the same reason the drawing reader is: a binary workbook
/// keeps its theme in the same XML part the modern one does.
/// </remarks>
internal static class ThemeLoader
{
    public static async ValueTask<WorkbookTheme> LoadAsync(
        OpcPackage package, string workbookPart, CancellationToken cancellationToken)
    {
        string? themePart = null;
        foreach (OpcRelationship relationship in
            await package.GetRelationshipsAsync(workbookPart, cancellationToken).ConfigureAwait(false))
        {
            if (relationship.Type == XlsxSchema.RelTheme && !relationship.IsExternal)
                themePart = OpcPath.Resolve(workbookPart, relationship.Target);
        }

        if (themePart is null || !package.PartExists(themePart))
            return WorkbookTheme.Office;

        try
        {
            byte[] content = await package.ReadPartBytesAsync(themePart, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            return WorkbookTheme.Read(stream);
        }
        catch (System.Xml.XmlException)
        {
            // A theme that will not parse costs nothing to fall back on: every slot the file
            // failed to state is one the standard scheme answers for anyway.
            return WorkbookTheme.Office;
        }
    }
}
