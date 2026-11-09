# Streaming: Constant Memory at Any Scale

The L1 layer does not build a document in memory: the writer serializes rows straight into the
package, and the reader yields rows as parsing proceeds. The only thing that accumulates is the
shared strings table (and it can be disabled).

## Writing: XlsxWriter

```csharp
await using XlsxWriter writer = await XlsxWriter.CreateAsync("big.xlsx");
XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data", new SheetOptions
{
    FrozenRows = 1,
    ColumnWidths = new Dictionary<int, double> { [1] = 20 },
});

sheet.WriteRow("Id", "Name", "Created", "Active");        // params ReadOnlySpan<CellValue>
for (int i = 1; i <= 1_000_000; i++)
{
    sheet.WriteRow(i, "name-" + (i % 1000), DateTime.UtcNow, i % 2 == 0);
    await sheet.FlushIfNeededAsync();                      // cheap; flushes buffer at threshold
}
```

Rules:

- `WriteRow` writes to a buffer synchronously; call `FlushIfNeededAsync` in the loop (or use
  `WriteRowAsync`), otherwise the buffer grows without flushing.
- One active sheet: `BeginSheetAsync` automatically finishes the previous one (zip is written strictly
  forward).
- Blank values (`CellValue.Blank`) create gaps; a completely empty row advances the counter without
  writing an element. `SkipRows(n)` skips rows explicitly.
- `Merge(range)` accumulates merges and writes them when the sheet finishes.
- Styles: register via `writer.RegisterStyle(style)` and pass `StyleId` as a parallel span or one
  style per row:

```csharp
StyleId header = writer.RegisterStyle(CellStyle.Default.Bold().Fill(0xFFEEEEEE));
sheet.WriteRow(["Id", "Name"], header);
```

- `XlsxWriterOptions.UseInlineStrings = true` disables the shared strings table — memory stays
  strictly O(1) even on a stream of unique text (the file is larger).
- Dates/durations without an explicit style automatically get builtin formats 22 / 46; otherwise
  Excel would show bare numbers.
- Document properties (`docProps`) are always written, as Excel writes them: `Application`, creation
  and modification dates, sheet list. Set them through the DOM (`Workbook.Properties`, see
  [worksheet.md](worksheet.md)); the streaming writer uses defaults.

## Reading: XlsxReader

```csharp
await using XlsxReader reader = await XlsxReader.OpenAsync("big.xlsx");
Console.WriteLine(string.Join(", ", reader.Sheets.Select(s => s.Name)));

await foreach (RowData row in reader.ReadRowsAsync("Data"))
{
    // row.Cells — ReadOnlySpan<CellValue>; index 0 = column A
    double id = row.Cells[0].GetNumber();
}
```

Streaming avoids a worksheet-sized DOM; it does not make caller-controlled work unbounded.
`XlsxReaderOptions.Budget` applies the same input, package, XML, cell, string, media and embedding
ceilings as `Workbook.LoadAsync`. See [Loading untrusted spreadsheets](loading-untrusted-input.md).

Rules:

- `RowData` is a view over a reused buffer: process or copy (`ToArray()`) before the next iteration.
- Empty rows are skipped — use `RowData.RowNumber` for position.
- Serial numbers with date formats arrive as `CellValueKind.DateTime`; duration formats (`[h]:mm:ss`)
  as `TimeSpan`; the 1904 date system is honored.
- Recoverable problems do not throw: subscribe to `XlsxReaderOptions.OnWarning` — codes like
  `BadNumber`, `BadSharedStringIndex` will arrive.
- The reader is not thread-safe; enumerate sheets sequentially.

## CSV and TSV Streaming & DOM Import/Export

`CsvReader` and `CsvWriter` provide true incremental RFC 4180 CSV and TSV (`CsvOptions.TabSeparated`,
`CsvOptions.SemicolonSeparated`) streaming over an 8 KiB chunk buffer (`O(row width)` memory) with
reusable `RowData` buffers (compatible with `[ExcelRow]` source-generated mappers), `TimeSpan`
duration inference (`[h]:mm:ss` / `hh:mm:ss`), and `DocumentLoadBudget` enforcement (`MaxInputBytes`,
`MaxCells`, `MaxCellTextCharacters`, `MaxTotalCellTextCharacters`), alongside convenience helpers
on `Workbook` and `Worksheet`. `MaxInputBytes` counts bytes of the source stream whether or not it
is seekable. Text after a field's closing quote stays in the field, as in Excel; set
`CsvOptions.StrictQuotes` to refuse that and unclosed quotes with `InvalidDataException`:

```csharp
// Streaming CSV/TSV rows (constant buffer, works with [ExcelRow] T.ReadFrom(in row)):
await foreach (RowData row in CsvReader.ReadRowsAsync("data.tsv", CsvOptions.TabSeparated))
{
    string region = row.GetValueOrBlank(0).GetText();
    double amount = row.GetValueOrBlank(1).GetNumber();
}

// Streaming CSV/TSV writer:
await using (CsvWriter writer = CsvWriter.Create("out.csv"))
{
    await writer.WriteRowAsync(new[] { CellValue.Text("Region"), CellValue.Text("Amount") });
    await writer.WriteRowAsync(new[] { CellValue.Text("North"), CellValue.Number(1250.5) });
}

// DOM convenience helpers on Workbook and Worksheet:
Workbook workbook = await Workbook.LoadCsvAsync("input.csv");
string csvText = workbook.Sheets[0].ToCsv();
await workbook.Sheets[0].SaveCsvAsync("output.csv");
```
