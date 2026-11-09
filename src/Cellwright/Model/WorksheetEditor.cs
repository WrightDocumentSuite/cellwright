using Cellwright.IO;

namespace Cellwright.Model;

/// <summary>
/// Inserting and deleting rows and columns. The cells are the easy part: everything anchored to
/// them — merges, links, notes, rules, tables, page breaks — has to move too, and every formula
/// in the workbook has to be re-aimed.
/// </summary>
internal static class WorksheetEditor
{
    /// <summary>Moves the grid of a sheet and everything that points into it.</summary>
    /// <param name="sheet">Sheet being edited.</param>
    /// <param name="rows">Whether rows move; otherwise columns do.</param>
    /// <param name="index">First 1-based row or column the edit touches.</param>
    /// <param name="count">How many appear, or a negative count of how many go away.</param>
    public static void Shift(Worksheet sheet, bool rows, int index, int count)
    {
        int limit = rows ? SheetLimits.MaxRow : SheetLimits.MaxColumn;
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, limit);
        if (count == 0)
            return;
        if (count < 0 && index - count - 1 > limit)
            throw new ArgumentOutOfRangeException(nameof(count), count, "The deletion runs past the end of the sheet.");

        ShiftGrids(sheet, rows, index, count);
        ShiftLineProperties(sheet, rows, index, count);

        var shift = new ReferenceShift(sheet.Name, rows, index, count);
        ShiftAnchoredObjects(sheet, shift);
        ShiftPageBreaks(sheet, rows, index, count);
        RetargetFormulas(sheet.Workbook, shift);
        DropStaleMarkup(sheet, rows, index, count);

        // The engine rebuilds its graph from scratch, which is exactly right here: every
        // coordinate it recorded has moved. This is not a cell edit, and must not be read as
        // somebody typing into whichever spill happens to cover the insertion point.
        sheet.Workbook.NotifyStructureChanged();
    }

    private static void ShiftGrids(Worksheet sheet, bool rows, int index, int count)
    {
        if (rows)
        {
            sheet.Values.ShiftRows(index, count);
            sheet.StyleIndexes.ShiftRows(index, count);
            sheet.Formulas.ShiftRows(index, count);
            return;
        }

        sheet.Values.ShiftColumns(index, count);
        sheet.StyleIndexes.ShiftColumns(index, count);
        sheet.Formulas.ShiftColumns(index, count);
    }

    /// <summary>
    /// Moves the height, width, visibility and outline settings. An inserted line takes the
    /// size and outline level of the one above or to its left, the way Excel formats a new row
    /// like its predecessor — but never its visibility, since a row that appeared invisible
    /// would be baffling.
    /// </summary>
    private static void ShiftLineProperties(Worksheet sheet, bool rows, int index, int count)
    {
        if (rows)
        {
            RowProperties above = sheet.RowTable.GetValueOrDefault(index - 1);
            Reindex(sheet.RowTable, index, count, above with { Hidden = false, Collapsed = false });
            return;
        }

        ColumnProperties left = sheet.ColumnTable.GetValueOrDefault(index - 1);
        Reindex(sheet.ColumnTable, index, count, left with { Hidden = false, Collapsed = false });
    }

    private static void Reindex<T>(Dictionary<int, T> table, int index, int count, T inherited)
        where T : struct
    {
        if (table.Count == 0 && count < 0)
            return;

        var moved = new Dictionary<int, T>(table.Count);
        foreach ((int line, T properties) in table)
        {
            if (line < index)
            {
                moved[line] = properties;
                continue;
            }

            if (count < 0 && line < index - count)
                continue;

            moved[line + count] = properties;
        }

        if (count > 0 && !EqualityComparer<T>.Default.Equals(inherited, default))
        {
            for (int line = index; line < index + count; line++)
                moved[line] = inherited;
        }

        table.Clear();
        foreach ((int line, T properties) in moved)
            table[line] = properties;
    }

    /// <summary>Moves everything the sheet anchors to a range: merges, links, notes, rules, tables.</summary>
    private static void ShiftAnchoredObjects(Worksheet sheet, ReferenceShift shift)
    {
        ShiftRanges(sheet.Merges, shift);

        sheet.Hyperlinks.Reanchor(range => shift.TryApply(ref range) ? range : null);
        sheet.Comments.Reanchor(cell => shift.TryApply(ref cell) ? cell : null);
        sheet.ThreadedComments.Reanchor(cell => shift.TryApply(ref cell) ? cell : null);
        sheet.Pictures.Reanchor(cell => shift.TryApply(ref cell) ? cell : null);
        sheet.Charts.Reanchor(cell => shift.TryApply(ref cell) ? cell : null);
        ReanchorMarkup(sheet, shift);

        foreach (DataValidation validation in sheet.Validations)
            validation.Ranges = MoveAll(validation.Ranges, shift);
        if (sheet.Validations.Count > 0)
            sheet.Validations.MarkModified();

        foreach (ConditionalFormatRule rule in sheet.ConditionalFormats)
        {
            rule.Ranges = MoveAll(rule.Ranges, shift);
            if (rule.Formulas.Count > 0)
                rule.Formulas = [.. rule.Formulas.Select(f => FormulaText.Shift(f, shift, sheet.Name)!)];
        }

        if (sheet.ConditionalFormats.Count > 0)
            sheet.ConditionalFormats.MarkModified();

        foreach (ProtectedRange range in sheet.Protection.EditableRanges)
            range.Ranges = MoveAll(range.Ranges, shift);

        ShiftTables(sheet, shift);
        ShiftSpills(sheet, shift);
        ShiftAutoFilter(sheet, shift);
        sheet.FormulaGroups.Clear();
    }

    /// <summary>
    /// Moves the preserved cell and formula attributes to the cells they still belong to.
    /// Leaving them behind would put a data table's description, or a metadata index, on
    /// whatever cell happened to slide into that position.
    /// </summary>
    private static void ReanchorMarkup(Worksheet sheet, ReferenceShift shift)
    {
        Rekey(sheet.FormulaMarkupTable, shift);
        Rekey(sheet.CellMarkupTable, shift);
        Rekey(sheet.RichStringTable, shift);

        static void Rekey<T>(Dictionary<CellRef, T> table, in ReferenceShift shift)
        {
            if (table.Count == 0)
                return;

            var moved = new Dictionary<CellRef, T>(table.Count);
            foreach ((CellRef original, T markup) in table)
            {
                CellRef cell = original;
                if (shift.TryApply(ref cell))
                    moved[cell] = markup;
            }

            table.Clear();
            foreach ((CellRef cell, T markup) in moved)
                table[cell] = markup;
        }
    }

    private static void ShiftTables(Worksheet sheet, in ReferenceShift shift)
    {
        bool changed = false;
        foreach (ExcelTable table in sheet.Tables.ToArray())
        {
            RangeRef range = table.Range;
            if (shift.TryApply(ref range) && range.Start.Row != range.End.Row)
            {
                changed |= range != table.Range;
                table.Reshape(range, ColumnsAfterShift(sheet, table, range, shift));
                continue;
            }

            // A table without its header row, or with nothing left of it, is not a table.
            sheet.Tables.Remove(table.Name);
            changed = true;
        }

        if (changed)
            sheet.Tables.MarkModified();
    }

    /// <summary>
    /// The column definitions for a table's new shape. Rows and columns that only move it leave
    /// them alone; a column edit inside it adds or drops definitions, and Excel requires a new
    /// column to have a name in the header row that matches its definition.
    /// </summary>
    private static IReadOnlyList<string> ColumnsAfterShift(
        Worksheet sheet, ExcelTable table, RangeRef range, in ReferenceShift shift)
    {
        RangeRef original = table.Range;
        if (shift.Rows || range.ColumnCount == original.ColumnCount)
            return table.Columns;

        if (shift.Count < 0)
        {
            int removedEnd = shift.Index - shift.Count - 1;
            List<string> kept = [];
            for (int i = 0; i < table.Columns.Count; i++)
            {
                int column = original.Start.Column + i;
                if (column < shift.Index || column > removedEnd)
                    kept.Add(table.Columns[i]);
            }

            return kept;
        }

        List<string> columns = [.. table.Columns];
        var used = new HashSet<string>(columns, StringComparer.OrdinalIgnoreCase);
        int position = shift.Index - original.Start.Column;
        int suffix = 1;

        for (int i = 0; i < shift.Count; i++)
        {
            string name;
            do
            {
                name = $"Column{suffix++}";
            }
            while (!used.Add(name));

            columns.Insert(position + i, name);
            if (table.HasHeaderRow)
                sheet.Values.Set(range.Start.Row, shift.Index + i, CellValue.Text(name));
        }

        return columns;
    }

    /// <summary>
    /// Moves each spill with its anchor. An anchor that was deleted takes the array with it, so
    /// what remains of its rectangle goes back to blank rather than staying as orphaned values.
    /// </summary>
    private static void ShiftSpills(Worksheet sheet, in ReferenceShift shift)
    {
        if (sheet.SpillTable.Count == 0)
            return;

        KeyValuePair<CellRef, RangeRef>[] spills = [.. sheet.SpillTable];
        sheet.SpillTable.Clear();

        foreach ((CellRef originalAnchor, RangeRef originalRange) in spills)
        {
            CellRef anchor = originalAnchor;
            RangeRef range = originalRange;
            bool rangeSurvives = shift.TryApply(ref range);

            if (shift.TryApply(ref anchor))
            {
                if (rangeSurvives)
                    sheet.SpillTable[anchor] = range;
                continue;
            }

            if (!rangeSurvives)
                continue;

            for (int row = range.Start.Row; row <= range.End.Row; row++)
            {
                for (int column = range.Start.Column; column <= range.End.Column; column++)
                    sheet.Values.Set(row, column, CellValue.Blank);
            }
        }
    }

    private static void ShiftAutoFilter(Worksheet sheet, in ReferenceShift shift)
    {
        if (sheet.AutoFilterRange is not { } filter)
            return;

        bool active = sheet.AutoFilterIsActive;
        sheet.AutoFilterRange = shift.TryApply(ref filter) ? filter : null;
        sheet.AutoFilterIsActive = active && sheet.AutoFilterRange is not null;
    }

    private static void ShiftPageBreaks(Worksheet sheet, bool rows, int index, int count)
    {
        IReadOnlyList<int> breaks = rows ? sheet.PageSetup.RowBreaks : sheet.PageSetup.ColumnBreaks;
        if (breaks.Count == 0)
            return;

        List<int> moved = [];
        foreach (int at in breaks)
        {
            if (at < index)
                moved.Add(at);
            else if (count > 0 || at >= index - count)
                moved.Add(at + count);
        }

        sheet.PageSetup.ReplaceBreaks(rows ? moved : null, rows ? null : moved);
    }

    private static void ShiftRanges(List<RangeRef> ranges, in ReferenceShift shift)
    {
        for (int i = ranges.Count - 1; i >= 0; i--)
        {
            RangeRef range = ranges[i];
            if (shift.TryApply(ref range))
                ranges[i] = range;
            else
                ranges.RemoveAt(i);
        }
    }

    private static IReadOnlyList<RangeRef> MoveAll(IReadOnlyList<RangeRef> ranges, in ReferenceShift shift)
    {
        List<RangeRef> moved = [];
        foreach (RangeRef original in ranges)
        {
            RangeRef range = original;
            if (shift.TryApply(ref range))
                moved.Add(range);
        }

        // A rule with no ranges left would be invalid; keeping one collapsed cell is the least
        // surprising outcome, and the caller can drop the rule if it no longer makes sense.
        if (moved.Count == 0)
            moved.Add(new RangeRef(new CellRef(shift.Index, 1)));
        return moved;
    }

    /// <summary>Re-aims every formula of the workbook and every defined name at the moved cells.</summary>
    private static void RetargetFormulas(Workbook workbook, ReferenceShift shift)
    {
        foreach (Worksheet other in workbook.Sheets)
        {
            string ownSheet = other.Name;
            other.Formulas.Transform(formula => FormulaText.Shift(formula, shift, ownSheet));

            // x14 icon-set formulas cannot contain local-cell or bang references, so a grid edit
            // cannot retarget a rule on another sheet. The edited sheet's xm:sqref can move,
            // however, and its raw form cannot be rewritten through the model. Remove only that
            // sheet's x14 CF container; unrelated extensions and other sheets remain untouched.
            if (string.Equals(other.Name, shift.SheetName, StringComparison.OrdinalIgnoreCase) &&
                other.Markup?.RemoveExtensionConditionalFormats() is true)
            {
                string axis = shift.Rows ? "row" : "column";
                string operation = shift.Count > 0
                    ? $"inserting {shift.Count} {axis}(s)"
                    : $"deleting {-shift.Count} {axis}(s)";
                workbook.OnWarning?.Invoke(new WorkbookWarning(
                    "StructuralEditDroppedExtensionConditionalFormats",
                    $"'{other.Name}': {operation} at {shift.Index} on '{shift.SheetName}' removed x14 " +
                    "conditional formatting whose raw ranges or formulas cannot be retargeted safely."));
            }

            foreach (DataValidation validation in other.Validations)
            {
                validation.Formula1 = FormulaText.Shift(validation.Formula1, shift, ownSheet);
                validation.Formula2 = FormulaText.Shift(validation.Formula2, shift, ownSheet);
            }

            foreach (Chart chart in other.Charts)
            {
                ReadOnlyMemory<byte> rewritten = ChartPart.RetargetShift(
                    chart.RawXml, shift, ownSheet, out bool changed);
                if (!changed)
                    continue;

                chart.ReplaceRawXml(rewritten);
                if (chart.SourcePartPath is { } partPath && workbook.Preserved is { } preserved &&
                    preserved.Parts.ContainsKey(partPath))
                {
                    preserved.Parts[partPath] = rewritten.ToArray();
                }
                else
                {
                    other.Charts.MarkModified();
                }
            }
        }

        foreach (DefinedName name in workbook.Names)
            name.RefersTo = FormulaText.Shift(name.RefersTo, shift, shift.SheetName) ?? name.RefersTo;
    }

    /// <summary>
    /// Throws away the preserved markup whose ranges the edit invalidated. Sort state, ignored
    /// errors and cell watches all record coordinates the model cannot follow, and leaving them
    /// pointing at the wrong cells would be worse than losing them.
    /// </summary>
    private static void DropStaleMarkup(Worksheet sheet, bool rows, int index, int count)
    {
        if (sheet.Markup is not { } markup)
            return;

        bool dropped = !string.IsNullOrEmpty(markup.BeforeMerges) || !string.IsNullOrEmpty(markup.PrintSetup);
        markup.BeforeMerges = null;
        markup.PrintSetup = null;
        if (!dropped)
            return;

        string what = rows ? "row" : "column";
        string how = count > 0 ? $"inserting {count} {what}(s)" : $"deleting {-count} {what}(s)";
        sheet.Workbook.OnWarning?.Invoke(new WorkbookWarning(
            "StructuralEditDroppedMarkup",
            $"'{sheet.Name}': {how} at {index} discarded sort state and ignored-error markers, " +
            "which record cell ranges the model does not represent."));
    }
}
