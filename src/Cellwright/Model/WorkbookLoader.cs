using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Styles;

namespace Cellwright.Model;

/// <summary>
/// Fills a <see cref="Workbook"/> from a package using the streaming reader. Loading is
/// recovery-oriented: problems downgrade to <see cref="WorkbookWarning"/>s wherever the
/// data allows it, and everything lands in <see cref="Workbook.LoadDiagnostics"/>.
/// </summary>
internal static class WorkbookLoader
{
    public static async ValueTask LoadAsync(
        Workbook workbook, Stream stream, LoadOptions options, CancellationToken cancellationToken)
    {
        var diagnostics = new LoadDiagnostics();
        void Warn(WorkbookWarning warning)
        {
            diagnostics.Add(warning);
            options.OnWarning?.Invoke(warning);
        }

        var readerOptions = new XlsxReaderOptions
        {
            LeaveOpen = options.LeaveOpen,
            OnWarning = Warn,
            Limits = options.Limits,
        };
        XlsxReader reader = await XlsxReader.OpenAsync(stream, readerOptions, cancellationToken).ConfigureAwait(false);
        await using (reader.ConfigureAwait(false))
        {
            workbook.SetDateSystem(reader.DateSystem);
            workbook.Protection.LockStructure = reader.Protection.LockStructure;
            workbook.Protection.LockWindows = reader.Protection.LockWindows;
            workbook.Protection.Password = reader.Protection.Password;
            workbook.Properties.CopyFrom(reader.Properties);

            // Everything this library writes is transitional, and a package cannot be half of
            // each conformance class: a preserved Strict chart next to a transitional workbook
            // is a file no consumer can make sense of. Reading is unaffected.
            bool preserve = options.PreserveUnknownParts && !reader.IsStrict;
            if (options.PreserveUnknownParts && reader.IsStrict)
            {
                Warn(new WorkbookWarning(
                    "StrictConformance",
                    "The workbook belongs to the Strict conformance class of ISO/IEC 29500. Its content reads " +
                    "normally, but saving produces a transitional package, so the parts this library does not " +
                    "model — charts, themes, drawings — are left behind rather than mixed in."));
            }

            workbook.Theme = await ThemeLoader.LoadAsync(
                reader.Package, reader.WorkbookPartPath, cancellationToken).ConfigureAwait(false);

            StylesPart styles = await reader.GetStylesAsync(cancellationToken).ConfigureAwait(false);
            var styleMap = new int[styles.CellFormats.Count];
            for (int i = 0; i < styleMap.Length; i++)
                styleMap[i] = workbook.StylePool.Intern(styles.CellFormats[i]);
            workbook.LoadedDifferentialStyles = styles.DifferentialStyles;

            foreach (SheetInfo info in reader.Sheets)
            {
                Worksheet sheet = AddSheetSafe(workbook, info.Name, Warn);
                sheet.Visibility = info.Visibility;
                sheet.SourceSheetId = info.SheetId;
                IReadOnlyList<OpcRelationship> relationships =
                    await reader.GetSheetRelationshipsAsync(info.Index, cancellationToken).ConfigureAwait(false);
                var sink = new WorksheetContentSink(sheet, styleMap, relationships, preserve, Warn);
                await reader.ReadSheetContentAsync(info.Index, sink, cancellationToken).ConfigureAwait(false);
                AttachConditionalStyles(sheet, styles.DifferentialStyles);
                await LoadCommentsAsync(sheet, reader, info.Index, relationships, Warn, cancellationToken).ConfigureAwait(false);
                await LoadThreadedCommentsAsync(sheet, reader, info.Index, relationships, Warn, cancellationToken).ConfigureAwait(false);
                await LoadTablesAsync(sheet, reader, info.Index, relationships, cancellationToken).ConfigureAwait(false);
                await DrawingLoader.LoadAsync(
                    reader.Package, sheet, reader.SheetPartPath(info.Index), relationships, Warn, cancellationToken)
                    .ConfigureAwait(false);
                if (preserve)
                    PreserveSheetLayout(sheet, reader, info.Index, relationships);
            }

            await LoadPeopleAsync(workbook, reader, Warn, cancellationToken).ConfigureAwait(false);
            await RecognizeDynamicArraysAsync(workbook, reader, cancellationToken).ConfigureAwait(false);
            LoadDefinedNames(workbook, reader.DefinedNames, Warn);
            workbook.ExternalBooks = await ExternalBookLoader.LoadAsync(reader, Warn, cancellationToken).ConfigureAwait(false);

            if (preserve)
                workbook.Preserved = await reader.CreatePreservationAsync(cancellationToken).ConfigureAwait(false);
        }

        workbook.AttachDiagnostics(diagnostics);
    }

    /// <summary>
    /// Shows each conditional rule the differential format it points at. The raw <c>dxfs</c>
    /// markup is still what gets written back; this parsed view is what lets code read — and
    /// edit — the paint of a rule that arrived from a file.
    /// </summary>
    private static void AttachConditionalStyles(Worksheet sheet, IReadOnlyList<DifferentialStyle> formats)
    {
        foreach (ConditionalFormatRule rule in sheet.ConditionalFormats)
        {
            if (!rule.IsOpaque && rule.Style is null &&
                rule.SourceStyleId is int index && (uint)index < (uint)formats.Count)
                rule.AttachSourceStyle(formats[index]);
        }
    }

    /// <summary>Reads the sheet's threaded comments, when it has a part for them.</summary>
    private static async ValueTask LoadThreadedCommentsAsync(
        Worksheet sheet,
        XlsxReader reader,
        int sheetIndex,
        IReadOnlyList<OpcRelationship> relationships,
        Action<WorkbookWarning> warn,
        CancellationToken cancellationToken)
    {
        string? part = null;
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type == XlsxSchema.RelThreadedComments && !relationship.IsExternal)
                part = OpcPath.Resolve(reader.SheetPartPath(sheetIndex), relationship.Target);
        }

        if (part is null || !reader.PartExists(part))
            return;

        try
        {
            byte[] content = await reader.ReadPartBytesAsync(part, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            ThreadedCommentsPart.Read(stream, sheet.ThreadedComments);
        }
        catch (System.Xml.XmlException exception)
        {
            warn(new WorkbookWarning(
                "BadThreadedComments",
                $"The threaded comments of sheet '{sheet.Name}' could not be read: {exception.Message}"));
        }
    }

    /// <summary>
    /// Tells the dynamic arrays apart from the legacy CSE formulas. Both are spelled
    /// <c>t="array"</c> with a <c>ref</c>; what separates them is the anchor's <c>cm</c> index
    /// pointing at an <c>XLDAPR</c> record in the metadata part.
    /// </summary>
    /// <remarks>
    /// A recognized anchor moves from the CSE groups to the sheet's spill table, cached
    /// rectangle included — which is what lets the first recalculation write over the spilled
    /// values the file already carries instead of mistaking them for somebody's cells.
    /// </remarks>
    private static async ValueTask RecognizeDynamicArraysAsync(
        Workbook workbook, XlsxReader reader, CancellationToken cancellationToken)
    {
        string? part = null;
        foreach (OpcRelationship relationship in await reader.GetWorkbookRelationshipsAsync(cancellationToken).ConfigureAwait(false))
        {
            if (relationship.Type == XlsxSchema.RelSheetMetadata && !relationship.IsExternal)
                part = OpcPath.Resolve(reader.WorkbookPartPath, relationship.Target);
        }

        if (part is null || !reader.PartExists(part))
            return;

        HashSet<int> dynamicIndexes;
        try
        {
            byte[] content = await reader.ReadPartBytesAsync(part, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            dynamicIndexes = MetadataPart.ReadDynamicArrayIndexes(stream);
        }
        catch (System.Xml.XmlException)
        {
            // A metadata part that will not parse leaves every array formula a CSE one, which
            // is also what a reader with no metadata support sees.
            return;
        }

        if (dynamicIndexes.Count == 0)
            return;

        foreach (Worksheet sheet in workbook.Sheets)
        {
            for (int i = sheet.FormulaGroups.Count - 1; i >= 0; i--)
            {
                FormulaGroup group = sheet.FormulaGroups[i];
                if (group.Kind != FormulaGroupKind.Array || !IsDynamic(sheet, group.Master, dynamicIndexes))
                    continue;

                sheet.SpillTable[group.Master] = group.Range;
                sheet.FormulaGroups.RemoveAt(i);
            }
        }

        static bool IsDynamic(Worksheet sheet, CellRef master, HashSet<int> dynamicIndexes)
        {
            if (!sheet.CellMarkupTable.TryGetValue(master, out List<KeyValuePair<string, string>>? attributes))
                return false;

            foreach ((string name, string value) in attributes)
            {
                if (name == "cm" &&
                    int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int index))
                    return dynamicIndexes.Contains(index);
            }

            return false;
        }
    }

    /// <summary>Reads the people who wrote the workbook's threaded comments.</summary>
    private static async ValueTask LoadPeopleAsync(
        Workbook workbook, XlsxReader reader, Action<WorkbookWarning> warn, CancellationToken cancellationToken)
    {
        string? part = null;
        foreach (OpcRelationship relationship in await reader.GetWorkbookRelationshipsAsync(cancellationToken).ConfigureAwait(false))
        {
            if (relationship.Type == XlsxSchema.RelPersons && !relationship.IsExternal)
                part = OpcPath.Resolve(reader.WorkbookPartPath, relationship.Target);
        }

        if (part is null || !reader.PartExists(part))
            return;

        try
        {
            byte[] content = await reader.ReadPartBytesAsync(part, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            ThreadedCommentsPart.ReadPeople(stream, workbook.People);
        }
        catch (System.Xml.XmlException exception)
        {
            warn(new WorkbookWarning(
                "BadThreadedComments",
                $"The list of comment authors could not be read: {exception.Message}"));
        }
    }

    /// <summary>Reads the sheet's notes from the comments part its relationships point at.</summary>
    private static async ValueTask LoadCommentsAsync(
        Worksheet sheet,
        XlsxReader reader,
        int sheetIndex,
        IReadOnlyList<OpcRelationship> relationships,
        Action<WorkbookWarning> warn,
        CancellationToken cancellationToken)
    {
        string? commentsPart = null;
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type == XlsxSchema.RelComments && !relationship.IsExternal)
                commentsPart = OpcPath.Resolve(reader.SheetPartPath(sheetIndex), relationship.Target);
        }

        if (commentsPart is null || !reader.PartExists(commentsPart))
            return;

        try
        {
            byte[] content = await reader.ReadPartBytesAsync(commentsPart, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            // The placeholder notes that stand in for threads are counted and dropped: they are
            // regenerated from the threads on save, so keeping them here would mean two models of
            // the same thing that could disagree.
            List<CellComment> notes = CommentsPart.Read(stream, out _);
            foreach (CellComment comment in notes)
                sheet.Comments.AddFromFile(comment);
        }
        catch (System.Xml.XmlException exception)
        {
            warn(new WorkbookWarning("BadComments", $"The notes of sheet '{sheet.Name}' could not be read: {exception.Message}"));
        }
    }

    /// <summary>Reads the tables the sheet's relationships point at.</summary>
    private static async ValueTask LoadTablesAsync(
        Worksheet sheet,
        XlsxReader reader,
        int sheetIndex,
        IReadOnlyList<OpcRelationship> relationships,
        CancellationToken cancellationToken)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type != XlsxSchema.RelTable || relationship.IsExternal)
                continue;

            string partPath = OpcPath.Resolve(reader.SheetPartPath(sheetIndex), relationship.Target);
            if (!reader.PartExists(partPath))
                continue;

            byte[] content = await reader.ReadPartBytesAsync(partPath, cancellationToken).ConfigureAwait(false);
            using var stream = new MemoryStream(content, writable: false);
            try
            {
                if (TablePart.Read(stream) is { } table)
                {
                    sheet.Tables.AddFromFile(table);
                    if (table.FilterIsActive)
                        sheet.MarkFilteredArea(table.Range);
                }
            }
            catch (System.Xml.XmlException)
            {
                // A malformed table part still travels with the file; it just is not modelled.
            }
        }
    }

    /// <summary>
    /// Remembers where a sheet's part lived and what it pointed at, for a lossless save.
    /// Hyperlink relationships are left out: the model owns those and regenerates them.
    /// </summary>
    private static void PreserveSheetLayout(
        Worksheet sheet, XlsxReader reader, int sheetIndex, IReadOnlyList<OpcRelationship> relationships)
    {
        sheet.SourcePartPath = reader.SheetPartPath(sheetIndex);
        sheet.SourceRelationshipId = reader.SheetRelationshipId(sheetIndex);
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type != XlsxSchema.RelHyperlink)
                sheet.SourceRelationships.Add(relationship);
        }
    }

    private static void LoadDefinedNames(
        Workbook workbook, IReadOnlyList<DefinedNameInfo> names, Action<WorkbookWarning> warn)
    {
        foreach (DefinedNameInfo info in names)
        {
            Worksheet? scope = null;
            if (info.LocalSheetIndex is { } index)
            {
                if ((uint)index >= (uint)workbook.Sheets.Count)
                {
                    warn(new WorkbookWarning("BadNameScope", $"Defined name '{info.Name}' targets sheet index {index}, which does not exist."));
                    continue;
                }

                scope = workbook.Sheets[index];
            }

            // Built-in names (print areas, filter ranges) are kept too: they belong to the
            // file's settings, and dropping them silently changes how the sheet prints.
            DefinedName defined = workbook.Names.Add(info.Name, info.RefersTo, scope);
            defined.IsHidden = info.IsHidden;
            defined.Markup = info.Attributes;
        }
    }

    /// <summary>Adds a sheet, repairing invalid or duplicate names instead of failing the load.</summary>
    internal static Worksheet AddSheetSafe(Workbook workbook, string name, Action<WorkbookWarning> warn)
    {
        try
        {
            return workbook.Sheets.Add(name);
        }
        catch (ArgumentException)
        {
            string fallback = $"Sheet{workbook.Sheets.Count + 1}";
            for (int suffix = 2; workbook.Sheets.Find(fallback) is not null; suffix++)
                fallback = $"Sheet{workbook.Sheets.Count + suffix}";
            warn(new WorkbookWarning("SheetRenamed", $"Sheet name '{name}' is invalid or duplicated; renamed to '{fallback}'."));
            return workbook.Sheets.Add(fallback);
        }
    }
}
