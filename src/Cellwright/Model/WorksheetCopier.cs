using Cellwright.IO;

namespace Cellwright.Model;

/// <summary>
/// Duplicates a worksheet into a fresh one. Everything the model represents is copied by
/// value, never shared: a rule, a note or a link is a mutable object, and a structural edit on
/// one sheet would otherwise reach across into the other.
/// </summary>
/// <remarks>
/// The copy owns nothing of the source's package layout. It has no part path, no relationship
/// ids and none of the source worksheet markup, which is what makes it a new sheet rather than
/// a second name for the same one. Pictures and charts have model representations and are
/// rebuilt into a new drawing graph; shapes and other opaque content are reported, not silently
/// presented as copied. Explicit self references are retargeted by default.
/// </remarks>
internal static class WorksheetCopier
{
    public static void Fill(
        Worksheet source,
        Worksheet target,
        SheetCopyOptions options,
        Action<WorkbookWarning> warn)
    {
        ReportUnsupportedContent(source, target, warn);
        CopyCells(source, target, options);
        CopyLayout(source, target);
        CopyView(source, target);
        CopyAttachments(source, target, options, warn);
        CopyPrintAndProtection(source, target);

        if (options.RetargetSelfReferences)
            target.RetargetSheetName(source.Name, target.Name);
    }

    private static void CopyCells(Worksheet source, Worksheet target, SheetCopyOptions options)
    {
        foreach ((int row, int column, CellValue value) in source.Values.EnumerateOccupied())
            target.Values.Set(row, column, value);
        foreach ((int row, int column, int style) in source.StyleIndexes.EnumerateOccupied())
            target.StyleIndexes.Set(row, column, style);
        foreach ((int row, int column, string? formula) in source.Formulas.EnumerateOccupied())
            target.Formulas.Set(row, column, formula);

        // Members of a shared group carry no text of their own, so the group is what says what
        // they compute; an array group is what keeps the master's ref attribute on save.
        foreach (FormulaGroup group in source.FormulaGroups)
        {
            string text = options.RetargetSelfReferences
                ? FormulaText.RenameSheet(group.Text, source.Name, target.Name)!
                : group.Text;
            target.FormulaGroups.Add(new FormulaGroup(group.Kind, group.Master, group.Range, text));
        }

        foreach ((CellRef cell, FormulaMarkup markup) in source.FormulaMarkupTable)
            target.FormulaMarkupTable[cell] = new FormulaMarkup([.. markup.Attributes], markup.WritesOwnElement);
        foreach ((CellRef cell, List<KeyValuePair<string, string>> markup) in source.CellMarkupTable)
            target.CellMarkupTable[cell] = [.. markup];
        foreach ((CellRef cell, RichStringMarkup markup) in source.RichStringTable)
            target.RichStringTable[cell] = markup;
        foreach ((CellRef anchor, RangeRef range) in source.SpillTable)
            target.SpillTable[anchor] = range;
    }

    private static void CopyLayout(Worksheet source, Worksheet target)
    {
        foreach ((int row, RowProperties properties) in source.RowTable)
            target.RowTable[row] = properties;
        foreach ((int column, ColumnProperties properties) in source.ColumnTable)
            target.ColumnTable[column] = properties;
        target.Merges.AddRange(source.Merges);
    }

    private static void CopyView(Worksheet source, Worksheet target)
    {
        target.Visibility = source.Visibility;
        target.FrozenRows = source.FrozenRows;
        target.FrozenColumns = source.FrozenColumns;
        target.TabColor = source.TabColor;
        target.ShowGridLines = source.ShowGridLines;
        target.ZoomScale = source.ZoomScale;
        target.AutoFilterRange = source.AutoFilterRange;
        target.AutoFilterIsActive = source.AutoFilterIsActive;
    }

    private static void CopyAttachments(
        Worksheet source,
        Worksheet target,
        SheetCopyOptions options,
        Action<WorkbookWarning> warn)
    {
        foreach (Hyperlink link in source.Hyperlinks)
            target.Hyperlinks.Add(link.Clone());
        foreach (CellComment comment in source.Comments)
            target.Comments.Add(comment.Cell, comment.Text, comment.Author).RevisionUid = comment.RevisionUid;
        CopyThreadedComments(source, target);
        foreach (DataValidation validation in source.Validations)
            target.Validations.Add(validation.Clone());
        foreach (ConditionalFormatRule rule in source.ConditionalFormats)
            target.ConditionalFormats.Add(rule.Clone());
        foreach (ExcelTable table in source.Tables)
            target.Tables.Add(Rename(table, UniqueTableName(target.Workbook, table.Name)));
        foreach (Picture picture in source.Pictures)
            target.Pictures.Add(Clone(picture));
        foreach (Chart chart in source.Charts)
            target.Charts.Add(Clone(chart, source, target, options, warn));

        // This flag now means only shapes outside the picture/chart model. Carrying it onto the
        // copy lets the saver repeat the warning if callers ignore the immediate diagnostics.
        target.DrawingHasOtherShapes = source.DrawingHasOtherShapes;
    }

    /// <summary>
    /// Copies conversations without letting message identities cross between sheets. Authors are
    /// workbook-wide and therefore keep their person ids; messages and parent links are local to
    /// the new threaded-comments part and receive a fresh, internally consistent id graph.
    /// </summary>
    private static void CopyThreadedComments(Worksheet source, Worksheet target)
    {
        if (source.ThreadedComments.Count == 0)
            return;

        target.ThreadedComments.ExtensionListXml = source.ThreadedComments.ExtensionListXml;

        var occupiedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Worksheet sheet in target.Workbook.Sheets)
        {
            foreach (ThreadedComment comment in sheet.ThreadedComments)
                occupiedIds.Add(comment.Id);
        }

        var newIds = new Dictionary<ThreadedComment, string>();
        var parentIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (ThreadedComment comment in source.ThreadedComments)
        {
            string id;
            do
            {
                id = Person.NewId();
            }
            while (!occupiedIds.Add(id));

            newIds.Add(comment, id);
            // Match Threads' handling of malformed duplicate ids: a reply names the first message
            // with that id. Valid files simply take the ordinary one-to-one path.
            parentIds.TryAdd(comment.Id, id);
        }

        foreach (ThreadedComment comment in source.ThreadedComments)
        {
            target.ThreadedComments.AddFromFile(new ThreadedComment
            {
                Cell = comment.Cell,
                Text = comment.Text,
                PersonId = comment.PersonId,
                Id = newIds[comment],
                ParentId = comment.ParentId is { } parent && parentIds.TryGetValue(parent, out string? mapped)
                    ? mapped
                    : null,
                Created = comment.Created,
                Resolved = comment.Resolved,
                MentionsXml = comment.MentionsXml,
                ExtensionListXml = comment.ExtensionListXml,
            });
        }

        target.ThreadedComments.MarkModified();
    }

    /// <summary>
    /// Copies a picture. The bytes are shared rather than duplicated — they are read-only and
    /// the same image on two sheets is one image — but the anchor is not, since moving the copy
    /// must not move the original.
    /// </summary>
    private static Picture Clone(Picture picture) => new()
    {
        Content = picture.Content,
        Format = picture.Format,
        Name = picture.Name,
        Description = picture.Description,
        IsHidden = picture.IsHidden,
        Anchor = Clone(picture.Anchor),
    };

    /// <summary>
    /// Copies a chart's complete chartSpace markup, not merely its typed projection. A loaded
    /// chart keeps the source part path as provenance so the saver can clone its child
    /// relationships onto the newly allocated chart part.
    /// </summary>
    private static Chart Clone(
        Chart chart,
        Worksheet source,
        Worksheet target,
        SheetCopyOptions options,
        Action<WorkbookWarning> warn)
    {
        ReadOnlyMemory<byte> content = chart.RawXml;
        if (options.RetargetSelfReferences)
        {
            content = ChartPart.RetargetSheetName(
                content, source.Name, target.Name, out bool parsed, out _);
            if (!parsed && !content.IsEmpty)
            {
                warn(new WorkbookWarning(
                    "SheetCopyChartReferencesNotRetargeted",
                    $"Chart '{chart.Name ?? "(unnamed)"}' on the copy '{target.Name}' was copied, " +
                    "but its malformed chartSpace could not be inspected for references to the source sheet."));
            }
        }

        return new Chart
        {
            RawXml = content,
            SourcePartPath = chart.SourcePartPath,
            Anchor = Clone(chart.Anchor),
            Name = chart.Name,
        };
    }

    private static PictureAnchor Clone(PictureAnchor anchor) => new()
    {
        Kind = anchor.Kind,
        From = anchor.From,
        FromOffset = anchor.FromOffset,
        To = anchor.To,
        ToOffset = anchor.ToOffset,
        Extent = anchor.Extent,
        Position = anchor.Position,
        MoveBehavior = anchor.MoveBehavior,
    };

    /// <summary>
    /// Reports source-owned material as one operation-level warning. Relationship ids inside
    /// opaque worksheet XML cannot be rebound without understanding the element that owns each
    /// one, so cloning that XML would be more dangerous than omitting it explicitly.
    /// </summary>
    private static void ReportUnsupportedContent(
        Worksheet source, Worksheet target, Action<WorkbookWarning> warn)
    {
        var omitted = new List<string>();
        if (source.Markup?.HasCopySensitiveContent == true)
            omitted.Add("opaque worksheet markup");
        if (source.DrawingHasOtherShapes)
            omitted.Add("unmodelled drawing shapes or text boxes");

        string[] relationshipRoles = [.. source.SourceRelationships
            .Where(static relationship => !IsRegeneratedRelationship(relationship.Type))
            .Select(static relationship => RelationshipRole(relationship.Type))
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)];
        if (relationshipRoles.Length > 0)
            omitted.Add($"package relationships ({string.Join(", ", relationshipRoles)})");

        if (omitted.Count == 0)
            return;

        warn(new WorkbookWarning(
            "SheetCopyContentNotCopied",
            $"The copy '{target.Name}' of '{source.Name}' omits source-only content: " +
            $"{string.Join("; ", omitted)}. Modelled cells, pictures and charts were copied."));
    }

    private static bool IsRegeneratedRelationship(string type) => type is
        XlsxSchema.RelHyperlink or XlsxSchema.RelComments or XlsxSchema.RelThreadedComments or
        XlsxSchema.RelVmlDrawing or XlsxSchema.RelTable or XlsxSchema.RelDrawing;

    private static string RelationshipRole(string type)
    {
        int separator = type.LastIndexOf('/');
        return separator >= 0 && separator + 1 < type.Length ? type[(separator + 1)..] : type;
    }

    private static void CopyPrintAndProtection(Worksheet source, Worksheet target)
    {
        PageSetup print = source.PageSetup;
        target.PageSetup.Orientation = print.Orientation;
        target.PageSetup.PaperSize = print.PaperSize;
        target.PageSetup.Scale = print.Scale;
        target.PageSetup.FitToWidth = print.FitToWidth;
        target.PageSetup.FitToHeight = print.FitToHeight;
        target.PageSetup.AutoPageBreaks = print.AutoPageBreaks;
        target.PageSetup.Margins = print.Margins;
        target.PageSetup.Header = print.Header;
        target.PageSetup.Footer = print.Footer;
        target.PageSetup.PrintGridLines = print.PrintGridLines;
        target.PageSetup.PrintHeadings = print.PrintHeadings;
        target.PageSetup.CenterHorizontally = print.CenterHorizontally;
        target.PageSetup.CenterVertically = print.CenterVertically;
        target.PageSetup.BlackAndWhite = print.BlackAndWhite;
        target.PageSetup.Draft = print.Draft;
        target.PageSetup.FirstPageNumber = print.FirstPageNumber;
        target.PageSetup.Order = print.Order;
        target.PageSetup.DifferentOddEven = print.DifferentOddEven;
        target.PageSetup.DifferentFirstPage = print.DifferentFirstPage;
        target.PageSetup.ScaleHeaderWithDocument = print.ScaleHeaderWithDocument;
        target.PageSetup.AlignHeaderWithMargins = print.AlignHeaderWithMargins;
        target.PageSetup.ReplaceBreaks([.. print.RowBreaks], [.. print.ColumnBreaks]);

        SheetProtection protection = source.Protection;
        target.Protection.IsEnabled = protection.IsEnabled;
        target.Protection.Password = protection.Password;
        target.Protection.AllowSelectLockedCells = protection.AllowSelectLockedCells;
        target.Protection.AllowSelectUnlockedCells = protection.AllowSelectUnlockedCells;
        target.Protection.AllowFormatCells = protection.AllowFormatCells;
        target.Protection.AllowFormatColumns = protection.AllowFormatColumns;
        target.Protection.AllowFormatRows = protection.AllowFormatRows;
        target.Protection.AllowInsertColumns = protection.AllowInsertColumns;
        target.Protection.AllowInsertRows = protection.AllowInsertRows;
        target.Protection.AllowInsertHyperlinks = protection.AllowInsertHyperlinks;
        target.Protection.AllowDeleteColumns = protection.AllowDeleteColumns;
        target.Protection.AllowDeleteRows = protection.AllowDeleteRows;
        target.Protection.AllowSort = protection.AllowSort;
        target.Protection.AllowAutoFilter = protection.AllowAutoFilter;
        target.Protection.AllowPivotTables = protection.AllowPivotTables;
        target.Protection.AllowEditObjects = protection.AllowEditObjects;
        target.Protection.AllowEditScenarios = protection.AllowEditScenarios;
        foreach (ProtectedRange range in protection.EditableRanges)
            target.Protection.EditableRanges.Add(range);
    }

    /// <summary>A table name is unique across the whole workbook, so the copy needs its own.</summary>
    private static string UniqueTableName(Workbook workbook, string name)
    {
        for (int suffix = 2; ; suffix++)
        {
            string candidate = $"{name}{suffix}";
            if (workbook.LocateTable(candidate) is null)
                return candidate;
        }
    }

    private static ExcelTable Rename(ExcelTable table, string name) => new()
    {
        Name = name,
        Range = table.Range,
        Columns = [.. table.Columns],
        HasHeaderRow = table.HasHeaderRow,
        HasTotalsRow = table.HasTotalsRow,
        StyleName = table.StyleName,
        ShowRowStripes = table.ShowRowStripes,
        ShowColumnStripes = table.ShowColumnStripes,
        DisplayName = table.DisplayName is null ? null : name,
    };
}
