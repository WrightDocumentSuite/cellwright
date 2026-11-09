# Cellwright

[![CI](https://github.com/WrightDocumentSuite/cellwright/actions/workflows/ci.yml/badge.svg)](https://github.com/WrightDocumentSuite/cellwright/actions/workflows/ci.yml)

[Website](https://wrightsuite.com/) · [Documentation](https://github.com/WrightDocumentSuite/cellwright/tree/main/docs) · [NuGet](https://www.nuget.org/packages/Cellwright)

A modern .NET library for Excel spreadsheets (Apache-2.0). Layered like `System.Text.Json`: primitives →
streaming → DOM → typed mapping through a source generator, plus a separate package for the
formula engine.

## What it does

- **Constant-memory streaming**: `XlsxWriter` (forward-only, UTF-8 straight to the wire,
  `params ReadOnlySpan<CellValue>`) and `XlsxReader` (`IAsyncEnumerable<RowData>`).
- **DOM**: `Workbook`/`Worksheet` with data-oriented storage — column slices in 512-row pages,
  not one object per cell. Values, styles, formulas (round-tripped as text), merges, freeze
  panes, hidden sheets.
- **A complete worksheet**: hidden rows and columns, grouping, auto-filter, hyperlinks, notes and
  threaded comments, Excel tables, data validation, conditional formatting, page setup, sheet and
  workbook protection, tab colour and zoom.
- **Structural editing**: inserting and deleting rows and columns, moving everything anchored to
  them and retargeting the whole workbook's formulas; reordering and copying sheets.
- **Document properties**: `Title`, `Author`, `Company`, dates and custom fields — what Excel
  shows on its Info tab, which a workbook built from scratch used not to carry at all.
- **Preservation-oriented editing of foreign files**: pivot tables, themes and `.xlsm` macros
  travel into the saved file as they were while their owning structures stay intact; operations
  that must rebuild a sheet or drawing report opaque content they cannot safely reproduce.
- **Styles**: immutable records with free interning; fluent builders on C# 14 extension members:
  `CellStyle.Default.Bold().Fill(0xFF2278D4).NumberFormat("#,##0.00")`. All of §18.8 is covered:
  pattern and gradient fills, diagonals, text rotation, theme colours, cell protection, whole-row
  and whole-column styles; named cell styles and the `dxf` formats of a loaded file are kept and
  readable.
- **Rich text**: character formatting inside a cell — `cell.RichText` reads the runs a file
  carries, `cell.SetRichText(...)` writes new ones.
- **`CellValue`** — a 24-byte tagged union with no boxing; dates at tick precision, honest serial
  date conversion (the 1900 bug, the 1904 system).
- **Typed mapping** (`Cellwright.Mapping`): `[ExcelRow]` + an incremental source generator →
  `IRowSchema<T>` (static abstract members), zero reflection, Native AOT.
- **Formula engine** (`Cellwright.Formulas`): 468 functions — 352 of the 356 in §18.17.7 of
  ISO/IEC 29500 (all but four Japanese and Thai ones), including securities, engineering, complex
  numbers, Bessel, correlation and hypothesis tests — with incremental recalculation over a
  dependency graph, cycle detection, shared-formula expansion, defined names, structured
  references `Table[Column]`, 3-D ranges `Q1:Q4!A1`, references into closed workbooks from the
  `externalLink` cache, `LET`, `LAMBDA` (`MAP`, `REDUCE`, `SCAN`, `MAKEARRAY`, `BYROW`, `BYCOL`,
  `ISOMITTED`) and `TEXT`.
- **Dynamic array spill**: a formula returning an array occupies its rectangle, `A1#` references
  the whole of it, and anything in the way makes it `#SPILL!` rather than a partial answer.
- **Charts**: embedded charts are read (kind, title, series with their references and cached
  points) and created (`sheet.AddChart(ChartKind.Column, "D2:K17", "Data!$B$2:$B$10")`); an
  untouched chart's markup is written back byte for byte.
- **Legacy `.xls`** (`Cellwright.Xls`): reads BIFF8, BIFF5 and BIFF2–4 into the same `Workbook`,
  writes BIFF8 (notes through the Office Art layer, pictures, conditional formatting, data
  validation and page setup), decrypts XOR/RC4/RC4 CryptoAPI, decodes and encodes Ptg formulas.
  Converting xls → xlsx falls out of that for free — pictures included — and `XlsReader` streams
  rows when the model is not wanted.
- **Binary `.xlsb`** (`Cellwright.Xls`): reads BIFF12 records into the same `Workbook` — values,
  styles, rows, formulas with structured references, notes, data validation, conditional
  formatting, tables and pictures, encrypted workbooks included. There is no `.xlsb` writer: save
  to xlsx or `.xls`.
- **Pictures and theme**: images are read from all three formats and written to xlsx and `.xls`,
  with the bytes never re-encoded; `xl/theme/theme1.xml` is parsed, so a theme colour resolves to
  a concrete ARGB along with its tint — see [pictures.md](docs/pictures.md).
- **Password to open**: an encrypted `.xlsx` and `.xlsb` are the same MS-CFB container, so
  `Cellwright.Xls` reads them: both MS-OFFCRYPTO schemes, AES/DES/3DES/RC2 ciphers, all five
  hashes, plus an HMAC check over the package. It writes them too — `SaveEncryptedAsync` puts a
  workbook under agile AES-256/SHA-512, the way Excel does; NPOI and ExcelDataReader open the
  result.
- **Macros**: VBA sources are extracted from `.xlsm` and `.xls` — MS-OVBA decompression, module
  kinds, referenced libraries, protection state. A project password does not hide the source.
- **Markdown export**: `book.ToMarkdown()` produces a GitHub Flavored Markdown document — a sheet
  becomes a section with a table, values rendered as Excel renders them, merges, hyperlinks and a
  gallery of pictures; `ExportMarkdownAsync` writes `workbook.md` and `media/` into a directory. A
  pure model → text conversion with no binary dependency — see [markdown.md](docs/markdown.md).
- **PDF output** (`Cellwright.Pdf`): `book.SaveAsPdf("report.pdf")` prints a workbook the way
  Excel prints it — the grid with its styles and number formats, merges, wrapping and overflow,
  rotation, pictures on their anchors, conditional formatting worked out rather than ignored,
  charts drawn, two-dimensional pagination with scaling and `FitToWidth`, repeated titles, headers
  and footers with `&P`/`&N`/`&A`. Optionally with a structure tree (a sheet is a table, cells are
  `TD`/`TH`). Nothing is launched outside the process, and no font is substituted silently — every
  substitution is named in the diagnostics; see [pdf-export.md](docs/pdf-export.md). The same
  package also sizes columns and rows to their contents: `sheet.AutoFit()`.
- **Diagnostics instead of silent failure**: recovery-oriented loading, `WorkbookWarning` with
  codes, `LoadDiagnostics`; a file of the wrong format (`.ods`, `.docx`, an encrypted package) is
  named for what it is rather than crashing the parser.
- **A foreign file is hostile input**: the finite `DocumentLoadBudget` caps source bytes, package
  expansion, XML, cells, text and resources across XLSX/XLSB/XLS — a zip bomb gets a stable
  `DocumentLoadLimitException`, not an `OutOfMemoryException`.
- **Async-first**: package I/O on the asynchronous `ZipArchive` APIs of .NET 10.

## A quick look

```csharp
using Cellwright;

// DOM
var book = Workbook.Create();
Worksheet sheet = book.Sheets.Add("Sales");
sheet["A1"] = "Product";
sheet["B1"] = 1234.5;
sheet.Range("A1:B1").Style(s => s.Bold().Fill(0xFF2278D4));
sheet.Cell("B3").SetFormula("SUM(B1:B2)");
sheet.Column("A").SetWidth(24);
sheet.AutoFilter("A1:B3");
await book.SaveAsync("sales.xlsx");

// Editing a real report: charts, pivot tables and macros stay where they were
Workbook report = await Workbook.LoadAsync("report.xlsm");
report.Sheets[0]["B7"] = 42;
await report.SaveAsync("report-updated.xlsm");

// Streaming millions of rows
await using XlsxWriter writer = await XlsxWriter.CreateAsync("big.xlsx");
XlsxSheetWriter data = await writer.BeginSheetAsync("Data");
for (int i = 1; i <= 1_000_000; i++)
{
    data.WriteRow(i, "name-" + (i % 1000), DateTime.UtcNow, i % 2 == 0);
    await data.FlushIfNeededAsync();
}
```

```csharp
// Typed rows (Cellwright.Mapping, source generator)
[ExcelRow]
public partial record Order(
    [ExcelColumn("Order ID")] int Id,
    string Customer,
    [ExcelColumn(Format = "#,##0.00")] decimal Total,
    DateTime? ShippedAt);

await foreach (Order order in reader.ReadRowsAsync<Order>("Orders")) { }
await sheet.WriteRowsAsync(orders);
```

```csharp
// Formulas (Cellwright.Formulas)
sheet.Cell("D2").SetFormula("B2*C2*VLOOKUP(A2,Rates!$A$2:$B$9,2,FALSE)");
sheet.Cell("D3").SetFormula("LET(total,SUM(Sales[Amount]),TEXT(total,\"#,##0.00\"))");
sheet.Cell("D4").SetFormula("PMT(0.08/12,60,30000)");            // -608.29
sheet.Cell("D5").SetFormula("DSUM(Records!A1:E99,\"Profit\",Criteria!A1:B2)");
sheet.Cell("D6").SetFormula("SUM(FILTER(B2:B99,A2:A99=\"North\"))");
sheet.Cell("D7").SetFormula("NORM.DIST(42,40,1.5,TRUE)");        // 0.908789
sheet.Cell("F2").SetFormula("SORT(UNIQUE(A2:A99))");             // spills down F
CalculationSummary summary = book.Calculate();   // the first pass computes everything
sheet["B2"] = 500;
book.Calculate();                                 // afterwards, only what depends on it
CellValue total = book.Evaluate("SUM(Report!D2:D9)");
```

```csharp
// Structural editing: everything anchored to the cells moves with them,
// and the workbook's formulas are retargeted
sheet.InsertRows(2, 3);
sheet.DeleteColumns(5);
sheet.Name = "Q1 2026";                       // formulas, names and chart references follow the sheet
SheetCopyResult copy = book.Sheets.Copy("Q1 2026", "Q2 2026", new SheetCopyOptions());
// Modelled content, pictures, charts and self-references are copied; inspect copy.Warnings
// if the source also contained opaque worksheet markup or unsupported drawing objects.
book.Sheets.Move("Summary", 0);

book.Properties.Title = "Quarterly report";   // the Info tab in Excel
book.Properties.Custom["Department"] = CellValue.Text("Finance");
```

```csharp
// Legacy .xls (Cellwright.Xls): read into the same Workbook, saved wherever you like
Workbook legacy = await XlsWorkbook.LoadAsync("report-2003.xls");
await legacy.SaveAsync("report-2026.xlsx");            // conversion out of the box
await legacy.SaveXlsAsync("report-2003-updated.xls");  // or back into binary

Workbook any = await Spreadsheet.LoadAsync(uploaded);  // xlsx, xls or xlsb — decided by content

// Binary .xlsb is read by the same call
Workbook binary = await Spreadsheet.LoadAsync("report.xlsb");
await binary.SaveAsync("report.xlsx");

// A password to open — for a binary file as much as for a package
Workbook locked = await Spreadsheet.LoadAsync("locked.xlsx", new XlsLoadOptions { Password = "secret" });

// Macros: sources and referenced libraries from both formats
VbaProject? macros = await VbaProject.LoadAsync("report.xlsm");
foreach (VbaModule module in macros?.Modules ?? [])
    Console.WriteLine($"{module.Kind} {module.Name}: {module.Code.Length} characters");
foreach (VbaReference reference in macros?.References ?? [])
    Console.WriteLine($"{reference.Name}: {reference.Description}");

// Values only, without the model
await using XlsReader rows = await XlsReader.OpenAsync("export.xls");
await foreach (RowData row in rows.ReadRowsAsync(0))
    Console.WriteLine(row.GetValueOrBlank(0));
```

## Performance

Measured on 2026-11-09 with Cellwright 0.1.4, .NET SDK 10.0.403 and runtime 10.0.14 on a shared Xeon Gold 5412U server. Full confidence intervals, exact versions and source commit are in [docs/benchmarks.md](docs/benchmarks.md).

100,000 rows × 20 columns (2 million cells).

| Method | Write mean | Allocated | Read mean | Allocated |
| --- | ---: | ---: | ---: | ---: |
| CellwrightStreaming | 1.373 s | 32.299 MiB | 2.024 s | 626.729 MiB |
| CellwrightDom | 1.293 s | 78.412 MiB | 2.097 s | 674.461 MiB |
| ClosedXml | 4.159 s | 2,094.892 MiB | 5.283 s | 2,617.666 MiB |
| EpPlus | 5.981 s | unavailable | 2.410 s | 1,782.266 MiB |

Shared-host timings include background-load variation. Compare methods within this snapshot; the older i7-8700 tables describe a different environment.

## Documentation

- [Getting started](docs/getting-started.md)
- [Architecture](docs/architecture.md)
- [Streaming](docs/streaming.md)
- [Worksheet: rows, columns and everything around cells](docs/worksheet.md)
- [Styles](docs/styles.md)
- [Mapping](docs/mapping.md)
- [Formulas](docs/formulas.md)
- [Legacy .xls](docs/xls.md)
- [Binary .xlsb](docs/xlsb.md)
- [Pictures and theme colours](docs/pictures.md)
- [Markdown export](docs/markdown.md)
- [PDF output](docs/pdf-export.md)
- [Macros (VBA)](docs/vba.md)
- [Specification conformance](docs/conformance.md)
- [Benchmarks](docs/benchmarks.md)

## Building

Requires the .NET 10 SDK.

```bash
dotnet build Cellwright.slnx
dotnet test Cellwright.slnx
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks -- --filter * --inProcess
dotnet run --project samples/Cellwright.Samples
```

CI builds and runs the fast fixture suite on Linux and Windows for every push and pull request
(`.github/workflows/ci.yml`). A scheduled integration job downloads public workbook corpora pinned
by commit and SHA-256, then rejects incomplete runs and unexpected skips from the TRX results.

## Layout

```
src/Cellwright                    core: primitives, OPC, streaming, DOM, styles
src/Cellwright.Formulas           formula engine: parser, evaluator, recalculation
src/Cellwright.Pdf                PDF output on Inkwright: grid, pages, headers, AutoFit
src/Cellwright.Xls                binary formats: CFB, BIFF8 (.xls), BIFF12 (.xlsb), Ptg formulas,
                                  decryption, VBA extraction
src/Cellwright.Mapping            mapping attributes + runtime
src/Cellwright.Mapping.Generator  Roslyn incremental generator (packed into Mapping)
tests/                            xUnit suite (run twice, the second time under a culture
                                  with a comma for the decimal separator), including the OpenXml
                                  validator, interop with ClosedXML, and a pinned public corpus of
                                  856 .xls, 73 .xlsb, 1,080 .xlsx and 32 .xlsm files. It checks
                                  package fidelity, values, links, properties, pictures and themes
                                  against ExcelDataReader/NPOI, cross-format twins, formula
                                  recalculation against cached Excel values, MS-OI29500
                                  divergences, and VBA module/reference extraction
benchmarks/                       BenchmarkDotNet vs EPPlus 8 / ClosedXML / ExcelDataReader
samples/                          working examples of every layer
```

The untrusted-input contract is in
[loading-untrusted-input.md](docs/loading-untrusted-input.md). The deliberate limitations are listed in [architecture.md](docs/architecture.md),
[worksheet.md](docs/worksheet.md), [formulas.md](docs/formulas.md), [xls.md](docs/xls.md),
[xlsb.md](docs/xlsb.md), [pictures.md](docs/pictures.md), [markdown.md](docs/markdown.md) and
[vba.md](docs/vba.md). In short: charts can be read, created (`Column`, `Bar`, `Line`, `Pie`,
`Doughnut`, `Area`, `Scatter`, `Bubble`, `Radar`, and combined Cartesian series) and updated via
`SetTitle`/`SetData`/`RefreshCache`, while untouched loaded charts and pivot tables (which have no
API) travel into the saved xlsx untouched; there is no R1C1 input;
an external workbook is computed from its link cache rather than by opening the file;
macros are read but never written; package encryption is written but `.xls` encryption is not;
`.xlsb` is read but not written; binary `.xls` does not preserve the parts it cannot model, and
writes the conditional formatting rules the classic records can express but not the kinds that
live in `CF12` records; Markdown is only written, never read back.
