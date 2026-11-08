# Cellwright

A modern .NET 10 Excel library, layered like `System.Text.Json`: primitives → constant-memory
streaming → mutable DOM → source-generated typed mapping. Data-oriented cell storage — column
slices in 512-row pages, not an object per cell. Async-first I/O, zero reflection, Native AOT
and trimming safe. Apache-2.0 licensed.

```csharp
using Cellwright;

// Read
Workbook book = await Workbook.LoadAsync("report.xlsx");
Worksheet sheet = book.Sheets[0];
CellValue total = sheet["B10"];

// Edit
sheet["C2"] = 1234.5;
sheet.Cell("C2").SetStyle(CellStyle.Default.Bold().NumberFormat("#,##0.00"));
sheet.Cell("C3").SetFormula("SUM(C1:C2)");

// Write
await book.SaveAsync("report.xlsx");
```

## What the core package covers

- **Constant-memory streaming**: forward-only `XlsxWriter` (UTF-8 straight to the wire) and
  `XlsxReader` (`IAsyncEnumerable<RowData>`) for files of any size.
- **A full worksheet model**: values, styles, formulas (round-tripped as text), merged cells,
  freeze panes, hidden rows and columns, grouping, auto-filter, hyperlinks, notes and threaded
  comments, Excel tables, data validation, conditional formatting, page setup, protection,
  document properties, rich text runs, embedded charts (read and create), pictures.
- **Preservation-oriented editing of foreign files**: pivot tables, themes and macros travel
  through byte for byte while their owning structures stay intact; operations that must rebuild
  a sheet or drawing emit coded diagnostics for opaque content they cannot safely reproduce.
- **Styles per ISO/IEC 29500 §18.8**: immutable records with free interning, fluent builders,
  pattern and gradient fills, theme colors, cell protection.
- **Recovery-oriented loading**: recoverable problems become coded `WorkbookWarning`s instead
  of exceptions, and hostile input (zip bombs, wrong formats) is named rather than crashed on.
- **Markdown export**: `book.ToMarkdown()` renders sheets as GitHub Flavored Markdown tables.

## The Cellwright family

| Package | Adds |
| --- | --- |
| `Cellwright` | The core: primitives, streaming, DOM (this package) |
| `Cellwright.Formulas` | Calculation engine: 447 functions, incremental recalc, dynamic array spill |
| `Cellwright.Mapping` | Typed row mapping via a Roslyn source generator |
| `Cellwright.Xls` | Legacy `.xls` read/write, `.xlsb` read, encryption, VBA extraction |
| `Cellwright.Pdf` | Prints workbooks to PDF the way Excel prints them |

## Documentation

Guides, per-feature docs and the full list of deliberate limitations live in the repository:
<https://github.com/WrightDocumentSuite/cellwright>

Known deviations from the spec are documented. Operations that rebuild opaque package structures
report coded diagnostics when content cannot be preserved or modelled faithfully.
