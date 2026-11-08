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

            if (!IsMeasurableRow(sheet, options, row) || TryGetMerge(sheet, row, column, out _))
                continue;

            CellStyle style = measurer.StyleOf(row, column);

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
        List<(RangeRef Range, double Height)> merged = [];
        foreach ((int row, int column, CellValue value) in sheet.Values.EnumerateOccupied())
        {
            if (value.IsBlank || row > lastRow)
                continue;

            if (!IsMeasurableRow(sheet, options, row))
                continue;

            if (TryGetMerge(sheet, row, column, out RangeRef merge))
            {
                // Only the top-left cell is displayed by a merge. Values stored in cells it
                // covers must not influence the height of content nobody sees.
                if (!options.IncludeMergedCells || merge.Start.Row != row || merge.Start.Column != column)
                    continue;

                double width = WidthOf(sheet, merge, digit, columnPoints);
                if (width <= 0)
                    continue;

                double mergedHeight = measurer.HeightNeeded(value, row, column, width);
                if (merge.RowCount == 1)
                    tallest[row] = Math.Max(tallest.GetValueOrDefault(row), mergedHeight);
                else
                    merged.Add((merge, mergedHeight));
                continue;
            }

            double points = ColumnPoints(sheet, column, digit, columnPoints);
            double needed = measurer.HeightNeeded(value, row, column, points);
            tallest[row] = Math.Max(tallest.GetValueOrDefault(row), needed);
        }

        foreach ((int row, double height) in tallest)
        {
            sheet.Row(row).Height =
                Math.Clamp(height, options.MinHeight, Math.Min(options.MaxHeight, SheetLimits.MaxRowHeight));
        }

        // A merge across several rows owns their combined height, not one row's height. Preserve
        // an already sufficient layout and grow as few rows as necessary when the text needs more.
        foreach ((RangeRef range, double height) in merged)
            GrowMergedRows(sheet, range, height, options);
    }

    /// <summary>The effective width of one column, cached because rows usually reuse columns.</summary>
    private static double ColumnPoints(
        Worksheet sheet, int column, double digit, Dictionary<int, double> columnPoints)
    {
        if (columnPoints.TryGetValue(column, out double points))
            return points;

        points = SheetUnits.ColumnPoints(
            sheet.Column(column).Width ?? SheetUnits.DefaultColumnWidth, digit);
        columnPoints[column] = points;
        return points;
    }

    /// <summary>The visible width of all columns a merged cell spans.</summary>
    private static double WidthOf(
        Worksheet sheet,
        RangeRef range,
        double digit,
        Dictionary<int, double> columnPoints)
    {
        double width = 0;
        for (int column = range.Start.Column; column <= range.End.Column; column++)
        {
            if (!sheet.Column(column).Hidden)
                width += ColumnPoints(sheet, column, digit, columnPoints);
        }

        return width;
    }

    /// <summary>
    /// Gives a multi-row merge enough combined height without shrinking rows whose existing total
    /// is already sufficient. Growth starts at the anchor row and observes the per-row ceiling.
    /// </summary>
    private static void GrowMergedRows(
        Worksheet sheet, RangeRef range, double needed, AutoFitOptions options)
    {
        double current = 0;
        for (int row = range.Start.Row; row <= range.End.Row; row++)
        {
            SheetRow target = sheet.Row(row);
            if (options.IncludeHiddenRows || !target.Hidden)
                current += target.Height ?? SheetUnits.DefaultRowHeight;
        }

        double remaining = needed - current;
        if (remaining <= 0)
            return;

        double maximum = Math.Min(options.MaxHeight, SheetLimits.MaxRowHeight);
        for (int row = range.Start.Row; row <= range.End.Row && remaining > 0; row++)
        {
            SheetRow target = sheet.Row(row);
            if (!options.IncludeHiddenRows && target.Hidden)
                continue;

            double height = target.Height ?? SheetUnits.DefaultRowHeight;
            if (height >= maximum)
                continue;

            double grown = Math.Min(maximum, Math.Max(options.MinHeight, height + remaining));
            target.Height = grown;
            remaining -= grown - height;
        }
    }

    /// <summary>The last row auto-fitting looks at, honouring the sample limit.</summary>
    private static int LastMeasuredRow(Worksheet sheet, AutoFitOptions options)
    {
        int last = sheet.Values.MaxRow;
        return options.SampleRows is int sample && sample > 0 ? Math.Min(last, sample) : last;
    }

    /// <summary>
    /// Whether a cell's row participates in fitting.
    /// </summary>
    private static bool IsMeasurableRow(Worksheet sheet, AutoFitOptions options, int row) =>
        options.IncludeHiddenRows || !sheet.Row(row).Hidden;

    /// <summary>Finds the merge a cell lies inside, if any.</summary>
    /// <remarks>
    /// Scanned rather than indexed: a merge is a rectangle of any size the sheet allows, and one
    /// spanning every column would cost an index entry per column. Sheets carry few merges, and a
    /// sheet with none — the common case — never reaches the loop at all.
    /// </remarks>
    private static bool TryGetMerge(Worksheet sheet, int row, int column, out RangeRef found)
    {
        IReadOnlyList<RangeRef> merges = sheet.MergedRanges;
        foreach (RangeRef merge in merges)
        {
            if (row >= merge.Start.Row && row <= merge.End.Row &&
                column >= merge.Start.Column && column <= merge.End.Column)
            {
                found = merge;
                return true;
            }
        }

        found = default;
        return false;
    }
}
