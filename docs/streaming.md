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

Rules:

- `RowData` is a view over a reused buffer: process or copy (`ToArray()`) before the next iteration.
- Empty rows are skipped — use `RowData.RowNumber` for position.
- Serial numbers with date formats arrive as `CellValueKind.DateTime`; duration formats (`[h]:mm:ss`)
  as `TimeSpan`; the 1904 date system is honored.
- Recoverable problems do not throw: subscribe to `XlsxReaderOptions.OnWarning` — codes like
  `BadNumber`, `BadSharedStringIndex` will arrive.
- The reader is not thread-safe; enumerate sheets sequentially.
