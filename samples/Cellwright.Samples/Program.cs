using Cellwright;
using Cellwright.Mapping;
using Cellwright.Samples;

string outputDirectory = Path.Combine(Path.GetTempPath(), "cellwright-samples");
Directory.CreateDirectory(outputDirectory);
Console.WriteLine($"Writing sample workbooks to {outputDirectory}");

await StreamingWriteSample();
await TypedMappingSample();
await DomReportSample();
await LoadModifySaveSample();
await WorksheetFeaturesSample();
await WorkbookOrganizationSample();
await MarkdownExportSample();
Console.WriteLine("Formula engine:");
await FormulaSample.RunAsync(outputDirectory);
Console.WriteLine("Legacy .xls:");
await XlsSample.RunAsync(outputDirectory);

Console.WriteLine("Done.");
return;

// 1. Constant-memory streaming write: a million cells without a DOM.
async Task StreamingWriteSample()
{
    string path = Path.Combine(outputDirectory, "streaming.xlsx");
    await using XlsxWriter writer = await XlsxWriter.CreateAsync(path);
    XlsxSheetWriter sheet = await writer.BeginSheetAsync("Numbers", new SheetOptions { FrozenRows = 1 });

    sheet.WriteRow("N", "Square", "Timestamp", "Even");
    for (int i = 1; i <= 50_000; i++)
    {
        sheet.WriteRow(i, (double)i * i, DateTime.Today.AddMinutes(i), i % 2 == 0);
        await sheet.FlushIfNeededAsync();
    }

    Console.WriteLine($"  streaming.xlsx: 50k rows, current row {sheet.CurrentRow}");
}

// 2. Typed rows through the source-generated schema: zero reflection, AOT-friendly.
async Task TypedMappingSample()
{
    string path = Path.Combine(outputDirectory, "orders.xlsx");
    List<Order> orders =
    [
        new(1, "Анна", 1234.56m, DateTime.Today),
        new(2, "Björn", 99.90m, null),
        new(3, "Chloé", 0.49m, DateTime.Today.AddDays(-1)),
    ];

    await using (XlsxWriter writer = await XlsxWriter.CreateAsync(path))
    {
        XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders");
        await sheet.WriteRowsAsync(orders);
    }

    await using XlsxReader reader = await XlsxReader.OpenAsync(path);
    await foreach (Order order in reader.ReadRowsAsync<Order>("Orders"))
        Console.WriteLine($"  order {order.Id}: {order.Customer} = {order.Total}");
}

// 3. DOM + fluent styling via C# 14 extension members.
async Task DomReportSample()
{
    var book = Workbook.Create();
    Worksheet sheet = book.Sheets.Add("Отчёт");

    sheet["A1"] = "Продажи за июль";
    sheet.Range("A1:D1").Merge().Style(s => s.Bold().FontSize(14).Center().Fill(0xFF2278D4).FontColor(ExcelColor.FromRgb(255, 255, 255)));

    string[] products = ["Яблоки", "Груши", "Сливы"];
    for (int i = 0; i < products.Length; i++)
    {
        int row = i + 2;
        sheet[row, 1] = products[i];
        sheet[row, 2] = Random.Shared.Next(100, 900);
        sheet[row, 3] = Random.Shared.NextDouble() * 100;
        sheet.Cell(row, 4).SetFormula($"B{row}*C{row}");
    }

    sheet.Range("C2:D4").NumberFormat("#,##0.00");
    sheet.Cell("A5").SetValue("Итого").Bold();
    sheet.Cell("D5").SetFormula("SUM(D2:D4)").Bold();
    sheet.SetColumnWidth(1, 18);
    sheet.FrozenRows = 1;

    await book.SaveAsync(Path.Combine(outputDirectory, "report.xlsx"));
    Console.WriteLine("  report.xlsx: styled DOM report with formulas");
}

// 4. Load, inspect diagnostics, modify, save.
async Task LoadModifySaveSample()
{
    string source = Path.Combine(outputDirectory, "report.xlsx");
    Workbook book = await Workbook.LoadAsync(source, new LoadOptions
    {
        OnWarning = warning => Console.WriteLine($"  [warning] {warning}"),
    });

    Worksheet sheet = book.Sheets[0];
    sheet["A7"] = $"Обновлено {DateTime.Now:yyyy-MM-dd HH:mm}";
    Console.WriteLine($"  report.xlsx loaded: used range {sheet.UsedRange}, warnings: {book.LoadDiagnostics!.Warnings.Count}");

    await book.SaveAsync(Path.Combine(outputDirectory, "report-updated.xlsx"));
}

// 5. The rest of a worksheet: visibility, grouping, links, notes, tables, input rules.
async Task WorksheetFeaturesSample()
{
    var book = Workbook.Create();
    Worksheet sheet = book.Sheets.Add("Каталог");

    sheet["A1"] = "Артикул";
    sheet["B1"] = "Название";
    sheet["C1"] = "В наличии";
    for (int row = 2; row <= 9; row++)
    {
        sheet[row, 1] = $"SKU-{row - 1:000}";
        sheet[row, 2] = $"Товар {row - 1}";
        sheet[row, 3] = row % 2 == 0 ? "Да" : "Нет";
    }

    sheet.Column("B").SetWidth(28);
    sheet.Rows(6, 9).Group(collapse: true);
    sheet.TabColor = ExcelColor.FromRgb(0x22, 0x78, 0xD4);
    sheet.AutoFilter("A1:C9");
    sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com/sku-001", "Открыть карточку");
    sheet.Comments.Add("C2", "Остаток сверялся 24 июля", author: "Склад");
    sheet.Validations.Add(DataValidation.List("C2:C9", "\"Да,Нет\""));
    sheet.Tables.Add("Каталог", "A1:C9");

    await book.SaveAsync(Path.Combine(outputDirectory, "catalog.xlsx"));
    Console.WriteLine("  catalog.xlsx: grouping, auto filter, links, notes, table, validation");
}

// 6. Organizing the workbook itself: sheet order, duplicates, and what the file says about itself.
async Task WorkbookOrganizationSample()
{
    var book = Workbook.Create();
    Worksheet january = book.Sheets.Add("Январь");
    january["A1"] = "Выручка";
    january["B1"] = 125_000;
    january.Cell("B2").SetFormula("B1*0.2");
    january.Column("A").SetWidth(20);

    // The copy takes everything the model represents and lands right after its original.
    Worksheet february = book.Sheets.Copy(january, "Февраль");
    february["B1"] = 141_500;

    book.Sheets.Add("Итоги").Cell("B1").SetFormula("Январь!B1+Февраль!B1");
    book.Sheets.Move("Итоги", 0);

    book.Properties.Title = "Выручка по месяцам";
    book.Properties.Author = "Анна Смирнова";
    book.Properties.Company = "Contoso";
    book.Properties.Custom["Отдел"] = CellValue.Text("Финансы");
    book.Properties.Custom["Утверждено"] = CellValue.Boolean(false);

    string path = Path.Combine(outputDirectory, "months.xlsx");
    await book.SaveAsync(path);

    Workbook reloaded = await Workbook.LoadAsync(path);
    Console.WriteLine(
        $"  months.xlsx: {string.Join(", ", reloaded.Sheets.Select(s => s.Name))}; " +
        $"автор '{reloaded.Properties.Author}', отдел '{reloaded.Properties.Custom["Отдел"]}'");
}

// 7. The same model as a document: a GitHub Flavored Markdown file beside its media directory.
async Task MarkdownExportSample()
{
    var book = Workbook.Create();
    book.Properties.Title = "Выручка за квартал";
    Worksheet sheet = book.Sheets.Add("Продажи");

    sheet["A1"] = "Товар";
    sheet["B1"] = "Сумма";
    sheet["C1"] = "Карточка";
    sheet.Cell("B1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Right));

    string[] products = ["Яблоки", "Груши"];
    for (int i = 0; i < products.Length; i++)
    {
        int row = i + 2;
        sheet[row, 1] = products[i];
        sheet[row, 2] = 1234.5 * (i + 1);
        sheet[row, 3] = "открыть";
        sheet.Hyperlinks.Add(RangeRef.Parse($"C{row}"), $"https://example.com/{i + 1}");
    }

    sheet.Range("B2:B3").NumberFormat("#,##0.00");

    string directory = Path.Combine(outputDirectory, "markdown");
    await book.ExportMarkdownAsync(directory);

    MarkdownDocument document = book.ToMarkdown();
    Console.WriteLine(
        $"  markdown/workbook.md: {document.Text.Length} символов, картинок {document.Images.Count}");
}

namespace Cellwright.Samples
{
    /// <summary>Typed row: the schema below is implemented by the source generator at build time.</summary>
    [ExcelRow]
    public partial record Order(
        [ExcelColumn("Order ID")] int Id,
        string Customer,
        [ExcelColumn(Format = "#,##0.00")] decimal Total,
        DateTime? ShippedAt);
}
