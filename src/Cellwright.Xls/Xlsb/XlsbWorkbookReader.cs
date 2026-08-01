using Cellwright.IO;
using Cellwright.Model;
using Cellwright.Streaming;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads a binary workbook end to end: finds the parts through the package's relationships,
/// reads the tables they hold, then each sheet, and hands back the same <see cref="Workbook"/>
/// the xlsx reader would have produced.
/// </summary>
internal static class XlsbWorkbookReader
{
    public static async Task<Workbook> ReadAsync(
        OpcPackage package, string workbookPart, XlsLoadOptions options, CancellationToken cancellationToken)
    {
        Workbook workbook = Workbook.Create();
        var diagnostics = new LoadDiagnostics();
        void Warn(WorkbookWarning warning)
        {
            diagnostics.Add(warning);
            options.OnWarning?.Invoke(warning);
        }

        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        List<string?> sheetRelationships = XlsbWorkbookPart.Read(
            await package.ReadPartBytesAsync(workbookPart, cancellationToken).ConfigureAwait(false),
            globals);

        workbook.SetDateSystem(globals.DateSystem);
        workbook.Theme = await ThemeLoader.LoadAsync(package, workbookPart, cancellationToken).ConfigureAwait(false);

        IReadOnlyList<OpcRelationship> relationships =
            await package.GetRelationshipsAsync(workbookPart, cancellationToken).ConfigureAwait(false);

        // The parts are found through the relationships rather than by their usual names: the
        // format does not fix where they live, and a package written by another tool is free to
        // put them elsewhere.
        if (Find(relationships, XlsxSchema.RelSharedStrings, workbookPart) is { } strings && package.PartExists(strings))
            XlsbStringTable.Read(await package.ReadPartBytesAsync(strings, cancellationToken).ConfigureAwait(false), globals);

        if (Find(relationships, XlsxSchema.RelStyles, workbookPart) is { } styles && package.PartExists(styles))
            XlsbStyles.Read(await package.ReadPartBytesAsync(styles, cancellationToken).ConfigureAwait(false), globals, workbook.StylePool);

        var sheetNames = new List<string>(globals.Sheets.Count);
        foreach (XlsSheetEntry entry in globals.Sheets)
            sheetNames.Add(entry.Name);

        var decoder = new PtgDecoder(globals, sheetNames);
        await ReadSheetsAsync(workbook, package, globals, sheetRelationships, relationships, workbookPart, decoder, Warn, cancellationToken)
            .ConfigureAwait(false);

        DefinedNameLoader.Apply(workbook, globals, decoder, Warn);
        workbook.AttachDiagnostics(diagnostics);
        return workbook;
    }

    private static async Task ReadSheetsAsync(
        Workbook workbook, OpcPackage package, XlsGlobals globals,
        List<string?> sheetRelationships, IReadOnlyList<OpcRelationship> relationships, string workbookPart,
        PtgDecoder decoder, Action<WorkbookWarning> warn, CancellationToken cancellationToken)
    {
        for (int index = 0; index < globals.Sheets.Count; index++)
        {
            XlsSheetEntry entry = globals.Sheets[index];
            string? part = index < sheetRelationships.Count
                ? Resolve(relationships, sheetRelationships[index], workbookPart)
                : null;

            if (part is null || !package.PartExists(part))
            {
                warn(new WorkbookWarning(
                    "SkippedSheet",
                    $"'{entry.Name}' is a chart, macro or dialog sheet, which this library does not model; " +
                    "it is not in the workbook that was loaded."));
                continue;
            }

            Worksheet sheet = WorkbookLoader.AddSheetSafe(workbook, entry.Name, warn);
            sheet.Visibility = entry.Visibility;

            IReadOnlyList<OpcRelationship> sheetLinks =
                await package.GetRelationshipsAsync(part, cancellationToken).ConfigureAwait(false);

            var sink = new WorksheetContentSink(sheet, styleMap: [], relationships: sheetLinks, preserveMarkup: false, warn);
            new XlsbSheetParser(globals, decoder, index, warn)
                .Parse(await package.ReadPartBytesAsync(part, cancellationToken).ConfigureAwait(false), sink);

            await ReadSheetPartsAsync(package, sheet, sheetLinks, part, decoder, warn, cancellationToken)
                .ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Reads the parts that hang off a sheet rather than living inside it: the notes, the
    /// tables defined over its cells, and the drawing. All are found through the sheet's own
    /// relationships.
    /// </summary>
    /// <remarks>
    /// Notes and tables are binary records like the sheet itself, but the drawing is not: this
    /// format keeps its drawings in the same DrawingML the modern one uses, so that part goes
    /// to the shared reader rather than to anything here.
    /// </remarks>
    private static async Task ReadSheetPartsAsync(
        OpcPackage package, Worksheet sheet, IReadOnlyList<OpcRelationship> links, string sheetPart,
        PtgDecoder decoder, Action<WorkbookWarning> warn, CancellationToken cancellationToken)
    {
        await DrawingLoader.LoadAsync(package, sheet, sheetPart, links, warn, cancellationToken).ConfigureAwait(false);

        foreach (OpcRelationship link in links)
        {
            if (link.IsExternal)
                continue;

            string target = OpcPath.Resolve(sheetPart, link.Target);
            if (!package.PartExists(target) || !target.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                continue;

            if (link.Type == XlsxSchema.RelComments)
            {
                XlsbCommentsPart.Read(
                    await package.ReadPartBytesAsync(target, cancellationToken).ConfigureAwait(false), sheet, warn);
            }
            else if (link.Type == XlsxSchema.RelTable)
            {
                XlsbTablePart.Read(
                    await package.ReadPartBytesAsync(target, cancellationToken).ConfigureAwait(false),
                    sheet,
                    decoder.Tables);
            }
        }
    }

    private static string? Find(IReadOnlyList<OpcRelationship> relationships, string type, string workbookPart)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type == type && !relationship.IsExternal)
                return OpcPath.Resolve(workbookPart, relationship.Target);
        }

        return null;
    }

    private static string? Resolve(IReadOnlyList<OpcRelationship> relationships, string? id, string workbookPart)
    {
        if (id is null)
            return null;

        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Id == id && !relationship.IsExternal)
                return OpcPath.Resolve(workbookPart, relationship.Target);
        }

        return null;
    }
}
