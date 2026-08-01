using Cellwright.Formulas;
using Cellwright.Xls;

namespace Cellwright.Samples;

/// <summary>
/// The legacy binary format: writing one, reading it back, and converting it to xlsx. A .xls
/// read here is an ordinary workbook, so the conversion is just a different save call.
/// </summary>
public static class XlsSample
{
    /// <summary>Runs the sample, writing its files into the given directory.</summary>
    public static async Task RunAsync(string outputDirectory)
    {
        string legacyPath = Path.Combine(outputDirectory, "legacy.xls");

        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Продажи");

        sheet["A1"] = CellValue.Text("Товар");
        sheet["B1"] = CellValue.Text("Количество");
        sheet["C1"] = CellValue.Text("Сумма");
        sheet.Range("A1:C1").Style(style => style with
        {
            Font = style.Font with { Bold = true },
            Fill = FillStyle.Solid(ExcelColor.FromRgb(0x22, 0x78, 0xD4)),
        });

        for (int row = 2; row <= 8; row++)
        {
            sheet[row, 1] = CellValue.Text($"Товар {row - 1}");
            sheet[row, 2] = CellValue.Number(row * 3);
            sheet.Cell(row, 3).SetFormula($"B{row}*100");
        }

        sheet.Cell("C9").SetFormula("SUM(C2:C8)");
        sheet.Merge(RangeRef.Parse("A10:C10"));
        sheet.FrozenRows = 1;

        // Notes, input rules and print settings all reach the binary format too.
        sheet.Comments.Add("C9", "Проверено вручную", "Ревизор");
        sheet.Validations.Add(DataValidation.WholeBetween("B2:B8", 1, 100));
        sheet.PageSetup.Orientation = PageOrientation.Landscape;
        sheet.PageSetup.Header = new HeaderFooterText { Default = "&CОтчёт о продажах" };

        // A formula in a .xls carries the value it last produced, so the file has to be
        // calculated before it is written or every formula cell reads as empty.
        original.Calculate();

        var warnings = new List<WorkbookWarning>();
        await original.SaveXlsAsync(legacyPath, new XlsSaveOptions { OnWarning = warnings.Add });
        Console.WriteLine($"  legacy.xls: {new FileInfo(legacyPath).Length} bytes, {warnings.Count} warnings");

        // Reading it back gives an ordinary Workbook, so saving it as xlsx converts the file.
        Workbook reloaded = await XlsWorkbook.LoadAsync(legacyPath);
        Worksheet read = reloaded.Sheets[0];
        Console.WriteLine($"  read back: '{read.Name}', C9 = {read["C9"]}, formula = {read.Cell("C9").Formula}");

        string convertedPath = Path.Combine(outputDirectory, "converted.xlsx");
        await reloaded.SaveAsync(convertedPath);
        Console.WriteLine($"  converted.xlsx: {new FileInfo(convertedPath).Length} bytes");

        // When the format is not known up front, the signature decides rather than the name.
        // The same call opens a package locked with a password to open: Excel wraps one in the
        // very same container, so both arrive looking like a legacy workbook.
        Workbook detected = await Spreadsheet.LoadAsync(legacyPath, new XlsLoadOptions { Password = "секрет" });
        Console.WriteLine($"  auto-detected: {detected.Sheets.Count} sheet(s)");

        // And when only the values matter, the rows can be walked without a model at all.
        await using XlsReader rows = await XlsReader.OpenAsync(legacyPath);
        int counted = 0;
        await foreach (RowData row in rows.ReadRowsAsync(0))
            counted++;
        Console.WriteLine($"  streamed: {counted} rows, notes read back: {read.Comments.Count}");
    }
}
