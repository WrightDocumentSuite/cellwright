using Cellwright.IO;

namespace Cellwright.Streaming;

/// <summary>
/// Walks the package from a sheet to its pictures: the sheet points at a drawing part, and the
/// drawing part points at the images its anchors embed.
/// </summary>
/// <remarks>
/// Written against the package rather than against either reader, because both package formats
/// arrive here. The modern format and the binary one differ in how a sheet is spelled, not in
/// how its drawing hangs off it.
/// </remarks>
internal static class DrawingLoader
{
    /// <summary>
    /// Reads the pictures of one sheet, if it has a drawing at all.
    /// </summary>
    /// <param name="package">The open package.</param>
    /// <param name="sheet">Sheet to fill.</param>
    /// <param name="sheetPart">Absolute part name of the sheet, which its relationships are relative to.</param>
    /// <param name="relationships">The sheet's relationships, already read.</param>
    /// <param name="warn">Where recoverable problems go.</param>
    /// <param name="cancellationToken">Cancellation.</param>
    public static async ValueTask LoadAsync(
        OpcPackage package,
        Worksheet sheet,
        string sheetPart,
        IReadOnlyList<OpcRelationship> relationships,
        Action<WorkbookWarning> warn,
        CancellationToken cancellationToken)
    {
        if (Find(relationships, XlsxSchema.RelDrawing, sheetPart) is not { } drawingPart ||
            !package.PartExists(drawingPart))
        {
            return;
        }

        try
        {
            IReadOnlyList<OpcRelationship> links =
                await package.GetRelationshipsAsync(drawingPart, cancellationToken).ConfigureAwait(false);

            Dictionary<string, ReadOnlyMemory<byte>> images =
                await ReadImagesAsync(package, drawingPart, links, cancellationToken).ConfigureAwait(false);

            byte[] content = await package.ReadPartBytesAsync(drawingPart, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            DrawingPart.Result result = DrawingPart.Read(stream, images);

            foreach (Picture picture in result.Pictures)
                sheet.Pictures.AddFromFile(picture);

            foreach (DrawingPart.ChartFrame frame in result.ChartFrames)
            {
                if (ResolveChart(links, frame.RelationshipId, drawingPart) is not { } chartPart ||
                    !package.PartExists(chartPart))
                {
                    continue;
                }

                sheet.Charts.AddFromFile(new Chart
                {
                    RawXml = await package.ReadPartBytesAsync(chartPart, cancellationToken).ConfigureAwait(false),
                    SourcePartPath = chartPart,
                    Anchor = frame.Anchor,
                    Name = frame.Name,
                });
            }

            // Pictures and chart frames both have complete write paths. Only anchors the model
            // could not classify remain a fidelity risk when a drawing is regenerated.
            sheet.DrawingHasOtherShapes = result.OtherShapes > 0;
        }
        catch (System.Xml.XmlException exception)
        {
            // The part exists but no anchors could be classified. Treat it as wholly opaque so
            // a later copy or drawing edit cannot imply that everything was represented.
            sheet.DrawingHasOtherShapes = true;
            warn(new WorkbookWarning(
                "BadDrawing", $"The drawing of sheet '{sheet.Name}' could not be read: {exception.Message}"));
        }
    }

    /// <summary>
    /// Reads every image the drawing embeds, keyed by the relationship id its anchors use. They
    /// are read up front because an anchor names an image by id and a package makes no promise
    /// that ids appear in any order.
    /// </summary>
    private static async ValueTask<Dictionary<string, ReadOnlyMemory<byte>>> ReadImagesAsync(
        OpcPackage package, string drawingPart, IReadOnlyList<OpcRelationship> links,
        CancellationToken cancellationToken)
    {
        var images = new Dictionary<string, ReadOnlyMemory<byte>>(StringComparer.Ordinal);
        foreach (OpcRelationship link in links)
        {
            // An image the package only links to lives somewhere else entirely; there are no
            // bytes here to hand back, and the anchor is reported as an unmodelled shape.
            if (link.Type != XlsxSchema.RelImage || link.IsExternal)
                continue;

            string target = OpcPath.Resolve(drawingPart, link.Target);
            if (!package.PartExists(target))
                continue;

            images[link.Id] = await package.ReadPartBytesAsync(target, cancellationToken).ConfigureAwait(false);
        }

        return images;
    }

    /// <summary>The part a chart frame's relationship points at, or null.</summary>
    private static string? ResolveChart(IReadOnlyList<OpcRelationship> links, string id, string drawingPart)
    {
        foreach (OpcRelationship link in links)
        {
            if (link.Id == id && link.Type == XlsxSchema.RelChart && !link.IsExternal)
                return OpcPath.Resolve(drawingPart, link.Target);
        }

        return null;
    }

    private static string? Find(IReadOnlyList<OpcRelationship> relationships, string type, string owner)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type == type && !relationship.IsExternal)
                return OpcPath.Resolve(owner, relationship.Target);
        }

        return null;
    }
}
