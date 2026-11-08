using Cellwright;
using Cellwright.Formulas;
using Cellwright.Mapping;
using Cellwright.Pdf;
using Cellwright.Xls;

var workbook = Workbook.Create();
Worksheet sheet = workbook.Sheets.Add("Gate");
sheet["A1"] = 20;
sheet["A2"] = 22;
sheet.Cell("A3").SetFormula("SUM(A1:A2)");

CalculationSummary calculation = workbook.Calculate();
if (calculation.ErrorCells > 0 || sheet["A3"].GetNumber() != 42)
    throw new InvalidOperationException("The packaged formula engine did not calculate 42.");

using (var xls = new MemoryStream())
{
    await XlsWorkbook.SaveAsync(workbook, xls);
    if (xls.Length == 0)
        throw new InvalidOperationException("The packaged XLS writer produced no output.");
}

PdfExportResult rendered = PdfExporter.Render(workbook);
using (rendered.Document)
{
    if (rendered.PageCount == 0 || rendered.Document.ToArray().Length == 0)
        throw new InvalidOperationException("The packaged PDF renderer produced no output.");
}

using var xlsx = new MemoryStream();
await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
    xlsx,
    new XlsxWriterOptions { LeaveOpen = true }))
{
    XlsxSheetWriter mapped = await writer.BeginSheetAsync("Rows");
    await mapped.WriteRowsAsync([new PackageRow(42, "package")]);
}

if (xlsx.Length == 0)
    throw new InvalidOperationException("The packaged Mapping analyzer generated no usable schema.");

Console.WriteLine("Cellwright package consumer passed.");

[ExcelRow]
public partial record PackageRow(
    [ExcelColumn("Value")] int Value,
    string Source);
