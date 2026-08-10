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

        // The print area may deliberately run far beyond the cells that hold anything. Walk the
        // sparse value grid rather than every logical cell in that rectangle; the row and column
        // maps below also leave hidden lines out exactly as the old rectangular walk did.
        foreach ((int row, int column, CellValue value) in _sheet.Values.EnumerateOccupied(plan.Area))
        {
            int r = grid.RowIndex(row);
            int c = grid.ColumnIndex(column);
            if (r < 0 || c < 0 || grid.RowHeightIsStated(r))
                continue;

            double needed = measurer.HeightNeeded(value, row, column, grid.WidthAt(c));
            if (needed > 0)
                grid.GrowRow(r, needed);
        }
    }
}
