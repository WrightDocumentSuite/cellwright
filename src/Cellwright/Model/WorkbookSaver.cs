using Cellwright.IO;
using Cellwright.Streaming;

namespace Cellwright.Model;

/// <summary>
/// Streams a <see cref="Workbook"/> model through <see cref="XlsxWriter"/>. The workbook's
/// style pool is shared with the writer, so <see cref="StyleId"/> indexes stay valid.
/// </summary>
internal static class WorkbookSaver
{
    public static async ValueTask SaveAsync(
        Workbook workbook, Stream stream, SaveOptions options, CancellationToken cancellationToken)
    {
        var writerOptions = new XlsxWriterOptions
        {
            UseInlineStrings = options.UseInlineStrings,
            LeaveOpen = options.LeaveOpen,
            DateSystem = workbook.DateSystem,
            // A calculated workbook already carries fresh values, so Excel can trust them.
            FullCalculationOnLoad = !workbook.IsFullyCalculated,
        };

        Action<WorkbookWarning> Warn = options.OnWarning ?? (static _ => { });

        XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, writerOptions, workbook.StylePool, workbook.Preserved, cancellationToken).ConfigureAwait(false);
        await using (writer.ConfigureAwait(false))
        {
            writer.Protection = workbook.Protection;
            writer.Properties = workbook.Properties;
            ReserveSheetLayout(workbook, writer);
            bool writesMetadata = NeedsMetadataPart(workbook, Warn);
            int nextTableId = 1;
            int nextDrawingIndex = 1;
            foreach (Worksheet sheet in workbook.Sheets)
            {
                SheetPlan plan = PlanSheet(sheet, writer, Warn, ref nextTableId, ref nextDrawingIndex, writesMetadata);
                XlsxSheetWriter target = await writer.BeginSheetAsync(sheet.Name, plan.Options, cancellationToken)
                    .ConfigureAwait(false);
                await WriteRowsAsync(sheet, target, cancellationToken).ConfigureAwait(false);
                foreach (RangeRef merge in sheet.MergedRanges)
                    target.Merge(merge);

                // The package holds one open entry at a time, so the sheet's satellite parts
                // can only be written once the sheet part itself is finished.
                await target.DisposeAsync().ConfigureAwait(false);
                await WriteCommentsAsync(sheet, writer, plan, cancellationToken).ConfigureAwait(false);
                await WriteTablesAsync(writer, plan, cancellationToken).ConfigureAwait(false);
                await WriteDrawingAsync(writer, plan, cancellationToken).ConfigureAwait(false);
            }

            if (writesMetadata)
            {
                await writer.WritePartAsync(
                    XlsxSchema.PartMetadata,
                    XlsxSchema.ContentTypeSheetMetadata,
                    MetadataPart.Write,
                    cancellationToken).ConfigureAwait(false);

                writer.AddWorkbookPart(XlsxSchema.RelSheetMetadata, XlsxSchema.PartMetadata);
            }

            await WritePeopleAsync(workbook, writer, cancellationToken).ConfigureAwait(false);
            WriteDefinedNames(workbook, writer);
        }
    }

    /// <summary>Everything decided about a sheet before its part is opened.</summary>
    private readonly record struct SheetPlan(
        SheetOptions Options,
        string? CommentsPart,
        string? VmlPart,
        int DrawingIndex,
        List<(string PartPath, ExcelTable Table)>? Tables,
        DrawingPlan? Drawing,
        VmlLeftovers? VmlLeftovers,
        string? ThreadedCommentsPart);

    /// <summary>The drawing part of a sheet whose pictures or charts the model owns.</summary>
    private readonly record struct DrawingPlan(
        string PartPath,
        List<Picture> Pictures,
        List<string> ImageRelationshipIds,
        List<(string PartPath, Picture Picture)> Images,
        List<(string PartPath, Chart Chart, string RelationshipId)> Charts);

    private static async ValueTask WriteTablesAsync(
        XlsxWriter writer, SheetPlan plan, CancellationToken cancellationToken)
    {
        if (plan.Tables is null)
            return;

        foreach ((string partPath, ExcelTable table) in plan.Tables)
        {
            await writer.WritePartAsync(
                partPath, XlsxSchema.ContentTypeTable, xml => TablePart.Write(xml, table), cancellationToken)
                .ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Writes the sheet's drawing and the images it embeds. The images go first: the drawing's
    /// own relationships name them, and a relationship pointing at a part that is not there
    /// yet is only correct because a package is read whole rather than in order.
    /// </summary>
    private static async ValueTask WriteDrawingAsync(
        XlsxWriter writer, SheetPlan plan, CancellationToken cancellationToken)
    {
        if (plan.Drawing is not { } drawing)
            return;

        foreach ((string partPath, Picture picture) in drawing.Images)
        {
            await writer.WriteBinaryPartAsync(
                partPath, DrawingWriter.ContentType(picture.Format), picture.Content, cancellationToken)
                .ConfigureAwait(false);
        }

        foreach ((string partPath, Chart chart, _) in drawing.Charts)
        {
            await writer.WriteBinaryPartAsync(
                partPath, XlsxSchema.ContentTypeChart, chart.RawXml, cancellationToken)
                .ConfigureAwait(false);
            if (chart.SourcePartPath is { } sourcePartPath)
                writer.ClonePreservedRelationships(sourcePartPath, partPath);
        }

        var frames = new List<DrawingWriter.ChartFrame>(drawing.Charts.Count);
        foreach ((_, Chart chart, string relationshipId) in drawing.Charts)
            frames.Add(new DrawingWriter.ChartFrame(chart.Anchor, relationshipId, chart.Name));

        await writer.WritePartAsync(
            drawing.PartPath,
            XlsxSchema.ContentTypeDrawing,
            xml => DrawingWriter.Write(xml, drawing.Pictures, drawing.ImageRelationshipIds, frames),
            cancellationToken).ConfigureAwait(false);

        writer.AddRelationships(drawing.PartPath, BuildDrawingRelationships(drawing));
    }

    private static List<OpcRelationship> BuildDrawingRelationships(in DrawingPlan drawing)
    {
        var links = new List<OpcRelationship>(drawing.Images.Count + drawing.Charts.Count);
        for (int i = 0; i < drawing.Images.Count; i++)
        {
            links.Add(new OpcRelationship(
                drawing.ImageRelationshipIds[i],
                XlsxSchema.RelImage,
                OpcPath.MakeRelative(drawing.PartPath, drawing.Images[i].PartPath)));
        }

        foreach ((string partPath, _, string relationshipId) in drawing.Charts)
        {
            links.Add(new OpcRelationship(
                relationshipId, XlsxSchema.RelChart, OpcPath.MakeRelative(drawing.PartPath, partPath)));
        }

        return links;
    }

    private static async ValueTask WriteCommentsAsync(
        Worksheet sheet, XlsxWriter writer, SheetPlan plan, CancellationToken cancellationToken)
    {
        if (plan.ThreadedCommentsPart is { } threadedPart)
        {
            await writer.WritePartAsync(
                threadedPart,
                XlsxSchema.ContentTypeThreadedComments,
                xml => ThreadedCommentsPart.Write(xml, sheet.ThreadedComments),
                cancellationToken).ConfigureAwait(false);
        }

        if (plan.VmlPart is not { } vmlPart)
            return;

        // A thread also needs a note beside it, so that a reader which predates threads shows
        // something on the cell. Those notes are generated rather than kept: the thread is what
        // the model owns, and a stale placeholder would name a message that is no longer there.
        List<CellComment> comments = [.. sheet.Comments];
        comments.AddRange(ThreadedCommentsPart.Placeholders(sheet.ThreadedComments.Threads, sheet.Workbook.People));
        comments.Sort(static (left, right) => left.Cell.Row != right.Cell.Row
            ? left.Cell.Row.CompareTo(right.Cell.Row)
            : left.Cell.Column.CompareTo(right.Cell.Column));

        if (plan.CommentsPart is { } commentsPart)
        {
            await writer.WritePartAsync(
                commentsPart, XlsxSchema.ContentTypeComments, xml => CommentsPart.Write(xml, comments), cancellationToken)
                .ConfigureAwait(false);
        }

        await writer.WritePartAsync(
            vmlPart,
            XlsxSchema.ContentTypeVmlDrawing,
            xml => CommentsPart.WriteVml(xml, comments, plan.DrawingIndex, plan.VmlLeftovers),
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Claims the part paths and relationship ids of sheets that came from a file before any
    /// sheet is written, so that a sheet added later cannot be handed one of them.
    /// </summary>
    private static void ReserveSheetLayout(Workbook workbook, XlsxWriter writer)
    {
        List<string> partPaths = [];
        List<string> relationshipIds = [];
        foreach (Worksheet sheet in workbook.Sheets)
        {
            if (sheet.SourcePartPath is { } partPath)
                partPaths.Add(partPath);
            if (sheet.SourceRelationshipId is { } relationshipId)
                relationshipIds.Add(relationshipId);
        }

        writer.ReserveLayout(partPaths, relationshipIds);
    }

    /// <summary>
    /// Whether this save writes a metadata part of its own: there are spills to mark, and the
    /// loaded file did not bring a metadata part whose record numbering ours would collide with.
    /// </summary>
    /// <remarks>
    /// A workbook that arrived with metadata keeps it — the part replays byte for byte and the
    /// anchors keep their <c>cm</c> indexes into it. A spill computed after loading such a file
    /// has no index to claim, and renumbering somebody else's part is not this writer's to do:
    /// the anchor is written as a plain array formula and said out loud.
    /// </remarks>
    private static bool NeedsMetadataPart(Workbook workbook, Action<WorkbookWarning> warn)
    {
        bool hasSpills = false;
        foreach (Worksheet sheet in workbook.Sheets)
            hasSpills |= sheet.SpillTable.Count > 0;

        if (!hasSpills)
            return false;

        bool preservedMetadata = workbook.Preserved is { } preserved &&
            preserved.Parts.ContainsKey(XlsxSchema.PartMetadata);

        if (!preservedMetadata)
            return true;

        foreach (Worksheet sheet in workbook.Sheets)
        {
            foreach (CellRef anchor in sheet.SpillTable.Keys)
            {
                if (!HasCellIndex(sheet, anchor))
                {
                    warn(new WorkbookWarning(
                        "DynamicArrayAsLegacy",
                        $"'{sheet.Name}'!{anchor} spilled after this file was loaded, and the file already " +
                        "carries a metadata part this library does not renumber. The formula is written as a " +
                        "legacy array formula: it keeps its range but will not resize in Excel.",
                        anchor));
                }
            }
        }

        return false;

        static bool HasCellIndex(Worksheet sheet, CellRef anchor)
        {
            if (!sheet.CellMarkupTable.TryGetValue(anchor, out List<KeyValuePair<string, string>>? attributes))
                return false;

            foreach ((string name, _) in attributes)
            {
                if (name == "cm")
                    return true;
            }

            return false;
        }
    }

    private static SheetPlan PlanSheet(
        Worksheet sheet, XlsxWriter writer, Action<WorkbookWarning> warn, ref int nextTableId, ref int nextDrawingIndex,
        bool writesMetadata)
    {
        Dictionary<CellRef, RangeRef>? arrayRanges = null;

        // A spill anchor is written the way the format spells a dynamic array: an array formula
        // over its rectangle, plus — when this save owns the metadata part — the cm index that
        // says it is dynamic rather than CSE.
        Dictionary<CellRef, List<KeyValuePair<string, string>>>? cellMarkup =
            sheet.CellMarkupTable.Count > 0 ? sheet.CellMarkupTable : null;

        if (sheet.SpillTable.Count > 0)
        {
            foreach ((CellRef anchor, RangeRef range) in sheet.SpillTable)
                (arrayRanges ??= [])[anchor] = range;

            if (writesMetadata)
            {
                var merged = new Dictionary<CellRef, List<KeyValuePair<string, string>>>(sheet.CellMarkupTable);
                foreach (CellRef anchor in sheet.SpillTable.Keys)
                {
                    if (!merged.ContainsKey(anchor))
                        merged[anchor] = [new KeyValuePair<string, string>("cm", "1")];
                }

                cellMarkup = merged;
            }
        }

        foreach (FormulaGroup group in sheet.FormulaGroups)
        {
            if (group.Kind == FormulaGroupKind.Array)
                (arrayRanges ??= [])[group.Master] = group.Range;
        }

        string partPath = sheet.SourcePartPath ?? writer.ClaimPartPath(XlsxSchema.WorksheetPart);
        List<OpcRelationship> relationships = [.. sheet.SourceRelationships];
        List<WrittenHyperlink>? hyperlinks = BuildHyperlinks(sheet, relationships);
        (string? commentsPart, string? vmlPart, string? legacyDrawingId, VmlLeftovers? vmlLeftovers) =
            PlanComments(sheet, writer, partPath, relationships);
        string? threadedPart = PlanThreadedComments(sheet, writer, partPath, relationships);
        (List<(string, ExcelTable)>? tableParts, List<string>? tableIds) =
            PlanTables(sheet, writer, partPath, relationships, ref nextTableId);
        (DrawingPlan? drawing, string? drawingId) = PlanDrawing(sheet, writer, partPath, relationships, warn);
        // Binary inputs have modelled rules but no OOXML block to replay. In that case the
        // model is the only source of truth even though AddFromFile correctly left it unmodified.
        bool conditionalFormatsChanged = sheet.ConditionalFormats.IsModified ||
            sheet.Markup?.ConditionalFormats is null;

        var options = new SheetOptions
        {
            FrozenRows = sheet.FrozenRows,
            FrozenColumns = sheet.FrozenColumns,
            Columns = sheet.ColumnTable.Count > 0 ? sheet.ColumnTable : null,
            Rows = sheet.RowTable.Count > 0 ? sheet.RowTable : null,
            ArrayFormulaRanges = arrayRanges,
            FormulaMarkup = sheet.FormulaMarkupTable.Count > 0 ? sheet.FormulaMarkupTable : null,
            CellMarkup = cellMarkup,
            RichStrings = sheet.RichStringTable.Count > 0 ? sheet.RichStringTable : null,
            Visibility = sheet.Visibility,
            TabColor = sheet.TabColor,
            ShowGridLines = sheet.ShowGridLines,
            ZoomScale = sheet.ZoomScale,
            Hyperlinks = hyperlinks,
            AutoFilterRange = sheet.AutoFilterRange,
            Validations = sheet.Validations.Count > 0 ? sheet.Validations : null,
            ValidationsChanged = sheet.Validations.IsModified || sheet.Markup?.DataValidations is null,
            ConditionalFormats = sheet.ConditionalFormats.Count > 0 ? sheet.ConditionalFormats : null,
            ConditionalStyleIds = conditionalFormatsChanged ? PlanConditionalStyles(sheet, writer) : null,
            ConditionalFormatsChanged = conditionalFormatsChanged,
            PageSetup = sheet.PageSetup,
            Protection = sheet.Protection,
            PartPath = partPath,
            RelationshipId = sheet.SourceRelationshipId,
            SheetId = sheet.SourceSheetId,
            Relationships = relationships.Count > 0 ? relationships : null,
            Markup = sheet.Markup,
            CommentsChanged = Regenerates(sheet),
            LegacyDrawingRelationshipId = legacyDrawingId,
            TablesChanged = sheet.Tables.IsModified,
            TableRelationshipIds = tableIds,
            PicturesChanged = sheet.Pictures.IsModified,
            DrawingRelationshipId = drawingId,
        };

        // Only a sheet that writes a drawing of its own consumes a shape-id block.
        int drawingIndex = vmlPart is null ? 0 : nextDrawingIndex++;
        return new SheetPlan(
            options, commentsPart, vmlPart, drawingIndex, tableParts, drawing, vmlLeftovers, threadedPart);
    }

    /// <summary>
    /// Writes the list of people who wrote the workbook's threaded comments, when the model owns
    /// it. A workbook whose threads were never touched keeps the list it arrived with.
    /// </summary>
    private static async ValueTask WritePeopleAsync(
        Workbook workbook, XlsxWriter writer, CancellationToken cancellationToken)
    {
        bool regenerates = workbook.People.IsModified;
        foreach (Worksheet sheet in workbook.Sheets)
            regenerates |= sheet.ThreadedComments.IsModified;

        if (!regenerates)
            return;

        string? sourcePart = null;
        if (workbook.Preserved is { } preserved)
        {
            foreach (OpcRelationship relationship in preserved.RelationshipsOf(preserved.WorkbookPartPath))
            {
                if (relationship.Type == XlsxSchema.RelPersons && !relationship.IsExternal)
                {
                    string candidate = OpcPath.Resolve(preserved.WorkbookPartPath, relationship.Target);
                    if (preserved.Parts.ContainsKey(candidate))
                        sourcePart = candidate;
                    writer.ReleasePreservedPart(candidate);
                }
            }
        }

        if (workbook.People.Count == 0)
        {
            // Every thread is gone, so the list of their authors goes with it.
            writer.DropWorkbookPart(XlsxSchema.RelPersons);
            return;
        }

        string part = sourcePart ?? writer.ClaimPartPath(XlsxSchema.PersonsPart);
        await writer.WritePartAsync(
            part,
            XlsxSchema.ContentTypePersons,
            xml => ThreadedCommentsPart.WritePeople(xml, workbook.People),
            cancellationToken).ConfigureAwait(false);

        writer.AddWorkbookPart(XlsxSchema.RelPersons, part);
    }

    /// <summary>
    /// Decides what happens to the sheet's threaded comments: kept as they arrived, or written
    /// afresh from the model.
    /// </summary>
    /// <remarks>
    /// The threads and the notes beside them are one decision, not two. A thread's placeholder
    /// note lives in the comments part, so rewriting either means rewriting both — which is why
    /// this is settled together with <see cref="PlanComments"/> rather than on its own.
    /// </remarks>
    private static string? PlanThreadedComments(
        Worksheet sheet, XlsxWriter writer, string sheetPartPath, List<OpcRelationship> relationships)
    {
        if (!Regenerates(sheet))
            return null;

        ReleaseRelatedParts(
            writer, sheetPartPath, relationships, static relationship =>
                relationship.Type == XlsxSchema.RelThreadedComments);
        relationships.RemoveAll(static r => r.Type == XlsxSchema.RelThreadedComments);
        if (sheet.ThreadedComments.Count == 0 &&
            sheet.ThreadedComments.ExtensionListXml is not { Length: > 0 })
            return null;

        var ids = new RelationshipIdAllocator();
        ids.Reserve(relationships);

        string part = writer.ClaimPartPath(XlsxSchema.ThreadedCommentsPart);
        relationships.Add(new OpcRelationship(
            ids.Next(), XlsxSchema.RelThreadedComments, OpcPath.MakeRelative(sheetPartPath, part)));

        return part;
    }

    /// <summary>
    /// Whether a sheet's notes and threads are written from the model rather than carried over.
    /// </summary>
    /// <remarks>
    /// Either kind of edit rewrites both parts, and so does an edit to the workbook's people:
    /// a placeholder note carries the author's name, so a renamed person changes the notes too.
    /// </remarks>
    private static bool Regenerates(Worksheet sheet) =>
        sheet.Comments.IsModified || sheet.ThreadedComments.IsModified || sheet.Workbook.People.IsModified;

    /// <summary>
    /// Decides what happens to the sheet's drawing. An untouched sheet keeps its original
    /// drawing part, which carries across the charts, shapes and text boxes that shared it; a
    /// change to the picture or chart list means the part is written afresh from the model.
    /// Modelled pictures and charts survive; unmodelled shapes are reported before omission.
    /// </summary>
    private static (DrawingPlan? Plan, string? RelationshipId) PlanDrawing(
        Worksheet sheet, XlsxWriter writer, string sheetPartPath, List<OpcRelationship> relationships,
        Action<WorkbookWarning> warn)
    {
        if (!sheet.Pictures.IsModified && !sheet.Charts.IsModified)
            return (null, null);

        relationships.RemoveAll(static r => r.Type == XlsxSchema.RelDrawing);
        if (sheet.Pictures.Count == 0 && sheet.Charts.Count == 0)
            return (null, null);

        if (sheet.DrawingHasOtherShapes)
        {
            warn(new WorkbookWarning(
                "DrawingShapesNotWritten",
                $"The drawing of '{sheet.Name}' was edited and is written afresh from the model. " +
                "The shapes and text boxes that shared the original drawing are not in the saved file; " +
                "its pictures and charts are."));
        }

        var ids = new RelationshipIdAllocator();
        ids.Reserve(relationships);
        string drawingPart = writer.ClaimPartPath(XlsxSchema.DrawingPart);
        string drawingId = ids.Next();
        relationships.Add(new OpcRelationship(
            drawingId, XlsxSchema.RelDrawing, OpcPath.MakeRelative(sheetPartPath, drawingPart)));

        // The drawing's own relationships are numbered from scratch: they belong to a part that
        // did not exist a moment ago and share nothing with the sheet's.
        var children = new RelationshipIdAllocator();
        List<Picture> pictures = [.. sheet.Pictures];
        List<string> imageIds = [];
        List<(string, Picture)> imageParts = [];
        foreach (Picture picture in pictures)
        {
            imageIds.Add(children.Next());
            imageParts.Add((
                writer.ClaimPartPath(index => XlsxSchema.MediaPart(index, DrawingWriter.Extension(picture.Format))),
                picture));
        }

        List<(string, Chart, string)> chartParts = [];
        foreach (Chart chart in sheet.Charts)
            chartParts.Add((writer.ClaimPartPath(XlsxSchema.ChartPart), chart, children.Next()));

        return (new DrawingPlan(drawingPart, pictures, imageIds, imageParts, chartParts), drawingId);
    }

    /// <summary>
    /// Resolves the differential format of every conditional formatting rule. A rule the model
    /// built interns its style into the workbook's format table; a rule read from a file whose
    /// style was not touched keeps the index it already had, because the format it points at is
    /// replayed verbatim — richer than the parsed view of it.
    /// </summary>
    private static List<int> PlanConditionalStyles(Worksheet sheet, XlsxWriter writer)
    {
        IReadOnlyList<DifferentialStyle> loaded = sheet.Workbook.LoadedDifferentialStyles;
        bool preservedFormatsTravel = sheet.Workbook.Preserved is not null;
        var ids = new List<int>(sheet.ConditionalFormats.Count);
        foreach (ConditionalFormatRule rule in sheet.ConditionalFormats)
        {
            if (rule.Style is { IsEmpty: false } style)
            {
                bool unchanged = preservedFormatsTravel &&
                    rule.SourceStyleId is int source && (uint)source < (uint)loaded.Count &&
                    loaded[source] == style;
                ids.Add(unchanged ? rule.SourceStyleId!.Value : writer.DifferentialStyles.Intern(style));
            }
            else
            {
                // Without the preserved package the original dxfs table is not written, and a
                // source index would point into whatever the pool happens to mint.
                ids.Add(preservedFormatsTravel ? rule.SourceStyleId ?? -1 : -1);
            }
        }

        return ids;
    }

    /// <summary>
    /// Decides what happens to the sheet's tables. Untouched ones keep their original parts,
    /// which carries their filter state and calculated columns across; a change to the table
    /// list means regenerating every part of the sheet from the model.
    /// </summary>
    private static (List<(string, ExcelTable)>? Parts, List<string>? RelationshipIds) PlanTables(
        Worksheet sheet, XlsxWriter writer, string sheetPartPath, List<OpcRelationship> relationships, ref int nextTableId)
    {
        if (!sheet.Tables.IsModified)
        {
            foreach (ExcelTable table in sheet.Tables)
                nextTableId = Math.Max(nextTableId, table.Id + 1);
            return (null, null);
        }

        relationships.RemoveAll(static r => r.Type == XlsxSchema.RelTable);
        if (sheet.Tables.Count == 0)
            return (null, null);

        var ids = new RelationshipIdAllocator();
        ids.Reserve(relationships);

        List<(string, ExcelTable)> parts = [];
        List<string> relationshipIds = [];
        foreach (ExcelTable table in sheet.Tables)
        {
            table.Id = nextTableId++;
            string partPath = writer.ClaimPartPath(XlsxSchema.TablePart);
            string relationshipId = ids.Next();
            relationships.Add(new OpcRelationship(
                relationshipId, XlsxSchema.RelTable, OpcPath.MakeRelative(sheetPartPath, partPath)));
            relationshipIds.Add(relationshipId);
            parts.Add((partPath, table));
        }

        return (parts, relationshipIds);
    }

    /// <summary>
    /// Decides what happens to the sheet's notes. Untouched notes keep their original comment
    /// and VML parts, which preserves the position and shape of every note box; edited notes
    /// mean a regenerated pair, and the old parts fall out of the package with their
    /// relationships. Regenerating the VML costs any form controls that shared it.
    /// </summary>
    private static (string? Comments, string? Vml, string? LegacyDrawingId, VmlLeftovers? Leftovers) PlanComments(
        Worksheet sheet, XlsxWriter writer, string sheetPartPath, List<OpcRelationship> relationships)
    {
        if (!Regenerates(sheet))
            return (null, null, null, null);

        // The original part may hold more than note boxes: form controls share the VML with
        // the comments, and regenerating it must not throw them away.
        VmlLeftovers? leftovers = LoadVmlLeftovers(sheet, sheetPartPath);
        ReleaseRelatedParts(
            writer, sheetPartPath, relationships, static relationship =>
                relationship.Type is XlsxSchema.RelComments or XlsxSchema.RelVmlDrawing);
        relationships.RemoveAll(static r => r.Type is XlsxSchema.RelComments or XlsxSchema.RelVmlDrawing);

        // A thread needs a note beside it, so the note count is both kinds together.
        int notes = sheet.Comments.Count + sheet.ThreadedComments.Threads.Count;
        if (notes == 0 && leftovers is null)
            return (null, null, null, null);

        var ids = new RelationshipIdAllocator();
        ids.Reserve(relationships);
        string? commentsPart = null;
        if (notes > 0)
        {
            commentsPart = writer.ClaimPartPath(XlsxSchema.CommentsPart);
            relationships.Add(new OpcRelationship(
                ids.Next(), XlsxSchema.RelComments, OpcPath.MakeRelative(sheetPartPath, commentsPart)));
        }

        string vmlPart = writer.ClaimPartPath(XlsxSchema.VmlDrawingPart);
        string legacyDrawingId = ids.Next();
        relationships.Add(new OpcRelationship(
            legacyDrawingId, XlsxSchema.RelVmlDrawing, OpcPath.MakeRelative(sheetPartPath, vmlPart)));

        return (commentsPart, vmlPart, legacyDrawingId, leftovers);
    }

    private static void ReleaseRelatedParts(
        XlsxWriter writer,
        string sourcePartPath,
        IReadOnlyList<OpcRelationship> relationships,
        Func<OpcRelationship, bool> owns)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (owns(relationship) && !relationship.IsExternal)
                writer.ReleasePreservedPart(OpcPath.Resolve(sourcePartPath, relationship.Target));
        }
    }

    /// <summary>
    /// Reads the non-note shapes out of the sheet's original VML part, when the package still
    /// has it to read.
    /// </summary>
    private static VmlLeftovers? LoadVmlLeftovers(Worksheet sheet, string sheetPartPath)
    {
        if (sheet.Workbook.Preserved is not { } preserved)
            return null;

        foreach (OpcRelationship relationship in sheet.SourceRelationships)
        {
            if (relationship.Type != XlsxSchema.RelVmlDrawing || relationship.IsExternal)
                continue;

            string vmlPath = OpcPath.Resolve(sheetPartPath, relationship.Target);
            return preserved.Parts.TryGetValue(vmlPath, out byte[]? content)
                ? VmlLeftovers.Parse(content)
                : null;
        }

        return null;
    }

    /// <summary>
    /// Turns the modelled hyperlinks into the shape the writer needs, allocating a sheet
    /// relationship for every external target alongside the ones the sheet already had.
    /// </summary>
    private static List<WrittenHyperlink>? BuildHyperlinks(Worksheet sheet, List<OpcRelationship> relationships)
    {
        if (sheet.Hyperlinks.Count == 0)
            return null;

        var ids = new RelationshipIdAllocator();
        ids.Reserve(relationships);

        List<WrittenHyperlink> written = [];
        foreach (Hyperlink link in sheet.Hyperlinks)
        {
            string? relationshipId = null;
            if (link.Target is { } target)
            {
                relationshipId = ids.Next();
                relationships.Add(new OpcRelationship(relationshipId, XlsxSchema.RelHyperlink, target, IsExternal: true));
            }

            written.Add(new WrittenHyperlink(link.Range, relationshipId, link.Location, link.Tooltip, link.Display));
        }

        return written;
    }

    private static void WriteDefinedNames(Workbook workbook, XlsxWriter writer)
    {
        foreach (DefinedName name in workbook.Names)
        {
            int? scopeIndex = null;
            if (name.Scope is { } scope)
            {
                int index = IndexOf(workbook, scope);
                if (index < 0)
                    continue; // The scope sheet was removed; the name goes with it.
                scopeIndex = index;
            }

            writer.AddDefinedName(name.Name, name.RefersTo, scopeIndex, name.IsHidden, name.Markup);
        }
    }

    private static int IndexOf(Workbook workbook, Worksheet sheet)
    {
        for (int i = 0; i < workbook.Sheets.Count; i++)
        {
            if (ReferenceEquals(workbook.Sheets[i], sheet))
                return i;
        }

        return -1;
    }

    private static async ValueTask WriteRowsAsync(Worksheet sheet, XlsxSheetWriter target, CancellationToken cancellationToken)
    {
        int maxRow = sheet.UsedRange?.End.Row ?? 0;
        foreach (int configured in sheet.RowTable.Keys)
            maxRow = Math.Max(maxRow, configured);
        if (maxRow == 0)
            return;

        int maxColumn = Math.Max(sheet.UsedRange?.End.Column ?? 1, 1);
        var values = new CellValue[maxColumn];
        var styles = new int[maxColumn];
        var formulas = new string?[maxColumn];

        for (int row = 1; row <= maxRow; row++)
        {
            bool occupied = sheet.Values.IsRowOccupied(row)
                || sheet.StyleIndexes.IsRowOccupied(row)
                || sheet.Formulas.IsRowOccupied(row)
                || sheet.RowTable.ContainsKey(row);
            if (!occupied)
                continue;

            sheet.Values.ReadRow(row, values, maxColumn);
            sheet.StyleIndexes.ReadRow(row, styles, maxColumn);
            sheet.Formulas.ReadRow(row, formulas, maxColumn);

            int count = maxColumn;
            while (count > 0 && values[count - 1].IsBlank && styles[count - 1] == 0 && formulas[count - 1] is null)
                count--;

            target.WriteRowAt(row, values.AsSpan(0, count), styles.AsSpan(0, count), formulas.AsSpan(0, count));
            await target.FlushIfNeededAsync(cancellationToken).ConfigureAwait(false);
        }
    }
}
