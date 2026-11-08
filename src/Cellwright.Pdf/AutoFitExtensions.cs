namespace Cellwright.Pdf;

/// <summary>
/// Sizing columns and rows to their contents, from the sheet.
/// </summary>
/// <remarks>
/// Auto-fitting needs to know how wide a font draws a string, which is why it arrives with the
/// printing package rather than the core. The measurements are the ones the printed page uses, so
/// a sheet fitted here prints without cutting the text it was fitted to.
/// </remarks>
/// <example>
/// <code>
/// Workbook book = await Workbook.LoadAsync("report.xlsx");
/// book.Sheets[0].AutoFit(new AutoFitOptions { MaxWidth = 40 });
/// await book.SaveAsync("report.xlsx");
/// </code>
/// </example>
public static class AutoFitExtensions
{
    extension(Worksheet sheet)
    {
        /// <summary>Sizes every column of the sheet to the widest cell in it.</summary>
        /// <param name="options">Limits and measuring settings, or <see langword="null"/> for the defaults.</param>
        /// <remarks>
        /// Cells that wrap their text are left out: those size their row instead. So are merged
        /// cells, whose text belongs to a rectangle rather than to any one column.
        /// </remarks>
        public Worksheet AutoFitColumns(AutoFitOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            options ??= new AutoFitOptions();

            using PdfExportContext context = PdfExportContext.ForMeasurement(sheet.Workbook, options.ToExportOptions());
            AutoFitter.Columns(context, sheet, options);
            return sheet;
        }

        /// <summary>Sizes every row of the sheet to the tallest cell in it.</summary>
        /// <param name="options">Limits and measuring settings, or <see langword="null"/> for the defaults.</param>
        /// <remarks>
        /// Unlike the measuring a print does, this sets the height of every row it measures,
        /// including rows that already state one: asking for it is asking for the rows to be
        /// resized. How tall a wrapped cell turns out depends on how wide its column is, so fit
        /// the columns first — <c>AutoFit</c> does both in that order. Merged cells are ignored
        /// unless <see cref="AutoFitOptions.IncludeMergedCells"/> is enabled.
        /// </remarks>
        public Worksheet AutoFitRows(AutoFitOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            options ??= new AutoFitOptions();

            using PdfExportContext context = PdfExportContext.ForMeasurement(sheet.Workbook, options.ToExportOptions());
            AutoFitter.Rows(context, sheet, options);
            return sheet;
        }

        /// <summary>Sizes the sheet's columns and then its rows.</summary>
        /// <param name="options">Limits and measuring settings, or <see langword="null"/> for the defaults.</param>
        public Worksheet AutoFit(AutoFitOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            options ??= new AutoFitOptions();

            using PdfExportContext context = PdfExportContext.ForMeasurement(sheet.Workbook, options.ToExportOptions());
            AutoFitter.Columns(context, sheet, options);
            AutoFitter.Rows(context, sheet, options);
            return sheet;
        }
    }

    extension(Workbook workbook)
    {
        /// <summary>Sizes the columns and rows of every sheet in the workbook.</summary>
        /// <param name="options">Limits and measuring settings, or <see langword="null"/> for the defaults.</param>
        /// <remarks>
        /// One measuring context serves the whole workbook, so a face shared by several sheets is
        /// resolved once rather than once per sheet.
        /// </remarks>
        public Workbook AutoFit(AutoFitOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            options ??= new AutoFitOptions();

            using PdfExportContext context = PdfExportContext.ForMeasurement(workbook, options.ToExportOptions());
            foreach (Worksheet sheet in workbook.Sheets)
            {
                AutoFitter.Columns(context, sheet, options);
                AutoFitter.Rows(context, sheet, options);
            }

            return workbook;
        }
    }

    extension(SheetColumn column)
    {
        /// <summary>Sizes this one column to the widest cell in it.</summary>
        /// <param name="options">Limits and measuring settings, or <see langword="null"/> for the defaults.</param>
        public SheetColumn AutoFit(AutoFitOptions? options = null)
        {
            options ??= new AutoFitOptions();
            Worksheet sheet = column.Sheet;

            using PdfExportContext context = PdfExportContext.ForMeasurement(sheet.Workbook, options.ToExportOptions());
            AutoFitter.Column(context, sheet, column.Number, options);
            return column;
        }
    }
}
