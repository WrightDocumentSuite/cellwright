using Cellwright.Pdf.Layout;
using Cellwright.Styles;

namespace Cellwright.Pdf;

/// <summary>
/// Sizing columns and rows to the text they hold.
/// </summary>
/// <remarks>
/// The measurements come from the same place the printed page takes them — real font metrics for
/// the face the style names, the number format rendered as it will be drawn, the same padding
/// around the glyphs. That is the whole reason auto-fitting lives in the printing package rather
/// than the core: a column can only be sized to its text by something that knows how wide the
/// text is, and nothing without a font does.
/// </remarks>
internal static class AutoFitter
{
    /// <summary>Sizes one column to the widest cell in it.</summary>
    public static void Column(PdfExportContext context, Worksheet sheet, int column, AutoFitOptions options) =>
        Columns(context, sheet, options, column);

    /// <summary>Sizes columns to the widest cell in each.</summary>
    /// <param name="context">The measuring context.</param>
    /// <param name="sheet">The sheet to fit.</param>
    /// <param name="options">Limits and measuring settings.</param>
    /// <param name="only">The one column to fit, or <see langword="null"/> for all of them.</param>
    public static void Columns(PdfExportContext context, Worksheet sheet, AutoFitOptions options, int? only = null)
    {
        var measurer = new CellMeasurer(context, sheet, new CellAppearances(context));
        double digit = SheetUnits.DigitPixels(measurer.Appearances.Standard.DigitWidth);
        int lastRow = LastMeasuredRow(sheet, options);

        Dictionary<int, double> widest = [];
        foreach ((int row, int column, CellValue value) in sheet.Values.EnumerateOccupied())
        {
            if (value.IsBlank || row > lastRow || (only is int wanted && column != wanted))
                continue;

            if (!IsMeasurable(sheet, options, row, column, out CellStyle style, measurer))
                continue;

            // A wrapped cell states how tall its row is, not how wide its column: sizing the
            // column to it would undo the wrapping the author asked for.
            if (style.Alignment.WrapText)
                continue;

            double needed = measurer.WidthNeeded(value, row, column);
            if (needed <= 0)
                continue;

            widest[column] = Math.Max(widest.GetValueOrDefault(column), needed);
        }

        foreach ((int column, double points) in widest)
        {
            SheetColumn target = sheet.Column(column);
            if (options.OnlySizedColumns && target.Width is null)
                continue;

            double width = SheetUnits.ColumnWidth(points, digit);
            target.Width = Math.Clamp(width, options.MinWidth, Math.Min(options.MaxWidth, SheetLimits.MaxColumnWidth));
        }
    }

    /// <summary>Sizes rows to the tallest cell in each, at the column widths now in effect.</summary>
    public static void Rows(PdfExportContext context, Worksheet sheet, AutoFitOptions options)
    {
        var measurer = new CellMeasurer(context, sheet, new CellAppearances(context));
        double digit = SheetUnits.DigitPixels(measurer.Appearances.Standard.DigitWidth);
        int lastRow = LastMeasuredRow(sheet, options);
        Dictionary<int, double> columnPoints = [];

        Dictionary<int, double> tallest = [];
        foreach ((int row, int column, CellValue value) in sheet.Values.EnumerateOccupied())
        {
            if (value.IsBlank || row > lastRow)
                continue;

            if (!IsMeasurable(sheet, options, row, column, out _, measurer))
                continue;

            if (!columnPoints.TryGetValue(column, out double points))
            {
                points = SheetUnits.ColumnPoints(
                    sheet.Column(column).Width ?? SheetUnits.DefaultColumnWidth, digit);
                columnPoints[column] = points;
            }

            double needed = measurer.HeightNeeded(value, row, column, points);
            tallest[row] = Math.Max(tallest.GetValueOrDefault(row), needed);
        }

        foreach ((int row, double height) in tallest)
        {
            sheet.Row(row).Height =
                Math.Clamp(height, options.MinHeight, Math.Min(options.MaxHeight, SheetLimits.MaxRowHeight));
        }
    }

    /// <summary>The last row auto-fitting looks at, honouring the sample limit.</summary>
    private static int LastMeasuredRow(Worksheet sheet, AutoFitOptions options)
    {
        int last = sheet.Values.MaxRow;
        return options.SampleRows is int sample && sample > 0 ? Math.Min(last, sample) : last;
    }

    /// <summary>
    /// Whether a cell has a say in the size of its row and column, and the style it wears.
    /// </summary>
    /// <remarks>
    /// A hidden row is not one the sheet is being sized for, and a merged cell is not either:
    /// its text belongs to a rectangle rather than to any one column, so Excel leaves merges out
    /// of auto-fitting and so does this.
    /// </remarks>
    private static bool IsMeasurable(
        Worksheet sheet, AutoFitOptions options, int row, int column, out CellStyle style, CellMeasurer measurer)
    {
        style = measurer.StyleOf(row, column);

        if (!options.IncludeHiddenRows && sheet.Row(row).Hidden)
            return false;

        return !IsMerged(sheet, row, column);
    }

    /// <summary>Whether a cell lies inside any of the sheet's merges.</summary>
    /// <remarks>
    /// Scanned rather than indexed: a merge is a rectangle of any size the sheet allows, and one
    /// spanning every column would cost an index entry per column. Sheets carry few merges, and a
    /// sheet with none — the common case — never reaches the loop at all.
    /// </remarks>
    private static bool IsMerged(Worksheet sheet, int row, int column)
    {
        IReadOnlyList<RangeRef> merges = sheet.MergedRanges;
        if (merges.Count == 0)
            return false;

        foreach (RangeRef merge in merges)
        {
            if (row >= merge.Start.Row && row <= merge.End.Row &&
                column >= merge.Start.Column && column <= merge.End.Column)
            {
                return true;
            }
        }

        return false;
    }
}
