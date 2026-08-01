namespace Cellwright.Pdf.Layout;

/// <summary>
/// The measuring a sheet needs before it can be cut into pages: how tall the rows that wrap
/// their text have to be.
/// </summary>
/// <remarks>
/// A row whose height the file states is left alone — Excel has already measured it, and its
/// figure is the one the author saw. Only a row that never had a height measures itself here,
/// which is also the only case where the answer can differ from Excel's by a font's metrics.
/// </remarks>
internal sealed partial class SheetComposer
{
    private void MeasureAutoHeights(SheetPlan plan, GridMeasurer grid)
    {
        var measurer = new CellMeasurer(_context, _sheet, grid.Appearances, _conditional);

        for (int r = 0; r < plan.Rows.Length; r++)
        {
            if (grid.RowHeightIsStated(r))
                continue;

            int row = plan.Rows[r];
            double needed = 0;

            for (int c = 0; c < plan.Columns.Length; c++)
            {
                int column = plan.Columns[c];
                CellValue value = _sheet.Values.Get(row, column);
                if (value.IsBlank)
                    continue;

                needed = Math.Max(needed, measurer.HeightNeeded(value, row, column, grid.WidthAt(c)));
            }

            if (needed > 0)
                grid.GrowRow(r, needed);
        }
    }
}
