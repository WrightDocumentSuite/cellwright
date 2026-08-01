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

        for (int row = group.Range.Start.Row; row <= group.Range.End.Row; row++)
        {
            for (int column = group.Range.Start.Column; column <= group.Range.End.Column; column++)
            {
                var cell = new CellRef(row, column);
                if (cell == group.Master || sheet.Formulas.Get(row, column) is not null)
                    continue;

                string text = FormulaRenderer.Render(
                    master, workbook, row - group.Master.Row, column - group.Master.Column);

                // Written straight into the slice: this restores what the file already meant,
                // so it must not look like a user edit that dirties the workbook.
                sheet.Formulas.Set(row, column, text);
            }
        }
    }
}
