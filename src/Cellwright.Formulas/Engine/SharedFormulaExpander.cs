using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas.Engine;

/// <summary>
/// Turns the shared formula groups the reader preserved into ordinary per-cell formulas by
/// re-rendering the master's program at each member's offset. Without this the members keep
/// their cached values but lose their formulas, which is what the core warns about.
/// </summary>
internal static class SharedFormulaExpander
{
    /// <summary>Expands every shared group of every sheet; array groups are left alone.</summary>
    public static void Expand(Workbook workbook, ProgramCache programs, Action<WorkbookWarning>? onWarning)
    {
        foreach (Worksheet sheet in workbook.Sheets)
        {
            if (sheet.FormulaGroups.Count == 0)
                continue;

            List<FormulaGroup> remaining = [];
            foreach (FormulaGroup group in sheet.FormulaGroups)
            {
                if (group.Kind == FormulaGroupKind.Shared)
                    ExpandGroup(sheet, group, workbook, programs, onWarning);
                else
                    remaining.Add(group);
            }

            sheet.FormulaGroups.Clear();
            sheet.FormulaGroups.AddRange(remaining);
        }
    }

    private static void ExpandGroup(
        Worksheet sheet, FormulaGroup group, Workbook workbook, ProgramCache programs, Action<WorkbookWarning>? onWarning)
    {
        FormulaProgram? master = programs.GetOrParse(group.Text, out FormulaParseException? failure);
        if (master is null)
        {
            onWarning?.Invoke(new WorkbookWarning("BadSharedFormula", failure!.Message, group.Master));
            return;
        }

        // Only the cells that declared membership: the range can enclose plain values, and
        // writing a formula into one of those would replace the value on the next calculation.
        foreach (CellRef cell in group.Members)
        {
            if (cell == group.Master || sheet.Formulas.Get(cell.Row, cell.Column) is not null)
                continue;

            string text = FormulaRenderer.Render(
                master, workbook, cell.Row - group.Master.Row, cell.Column - group.Master.Column);

            // Written straight into the slice: this restores what the file already meant,
            // so it must not look like a user edit that dirties the workbook.
            sheet.Formulas.Set(cell.Row, cell.Column, text);
        }
    }
}
