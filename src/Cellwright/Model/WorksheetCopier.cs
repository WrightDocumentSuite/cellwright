using Cellwright.IO;

namespace Cellwright.Model;

/// <summary>
/// Duplicates a worksheet into a fresh one. Everything the model represents is copied by
/// value, never shared: a rule, a note or a link is a mutable object, and a structural edit on
/// one sheet would otherwise reach across into the other.
/// </summary>
/// <remarks>
/// The copy owns nothing of the source's package layout. It has no part path, no relationship
/// ids and none of the source markup, which is what makes it a new sheet rather than a second
/// name for the same one — and what costs it the drawings and pictures the model does not
/// represent. Formulas travel as text, so references within the sheet follow the copy while
/// ones that name the source sheet outright keep pointing at it.
/// </remarks>
internal static class WorksheetCopier
{
    public static void Fill(Worksheet source, Worksheet target)
    {
        CopyCells(source, target);
        CopyLayout(source, target);
        CopyView(source, target);
        CopyAttachments(source, target);
        CopyPrintAndProtection(source, target);
    }

    private static void CopyCells(Worksheet source, Worksheet target)
    {
        foreach ((int row, int column, CellValue value) in source.Values.EnumerateOccupied())
            target.Values.Set(row, column, value);
        foreach ((int row, int column, int style) in source.StyleIndexes.EnumerateOccupied())
            target.StyleIndexes.Set(row, column, style);
        foreach ((int row, int column, string? formula) in source.Formulas.EnumerateOccupied())
            target.Formulas.Set(row, column, formula);

        // Members of a shared group carry no text of their own, so the group is what says what
        // they compute; an array group is what keeps the master's ref attribute on save.
        target.FormulaGroups.AddRange(source.FormulaGroups);

        foreach ((CellRef cell, FormulaMarkup markup) in source.FormulaMarkupTable)
            target.FormulaMarkupTable[cell] = markup;
        foreach ((CellRef cell, List<KeyValuePair<string, string>> markup) in source.CellMarkupTable)
            target.CellMarkupTable[cell] = [.. markup];
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

    private static void CopyAttachments(Worksheet source, Worksheet target)
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
        Anchor = new PictureAnchor
        {
            Kind = picture.Anchor.Kind,
            From = picture.Anchor.From,
            FromOffset = picture.Anchor.FromOffset,
            To = picture.Anchor.To,
            ToOffset = picture.Anchor.ToOffset,
            Extent = picture.Anchor.Extent,
            Position = picture.Anchor.Position,
            MoveBehavior = picture.Anchor.MoveBehavior,
        },
    };

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
