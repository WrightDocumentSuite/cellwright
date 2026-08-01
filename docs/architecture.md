# Cellwright Architecture

Cellwright is an xlsx library for .NET 10, layered like `System.Text.Json`: low-level streaming,
a document model on top, then typed mapping through a source generator. Each layer is public and
standalone.

## Layers

```
L3  Cellwright.Mapping   IRowSchema<T> + incremental source generator (zero reflection, AOT)
    Cellwright.Formulas  parser → RPN program, stack evaluator, dependency graph
    Cellwright.Pdf       print: grid → pages → content operators (on top of Inkwright)
    Cellwright.Xls       binary formats: CFB, BIFF8 (.xls), BIFF12 (.xlsb), Ptg, decryption
L2  DOM                  Workbook / Worksheet / Cell / CellRange — mutable in-memory model
L1  Streaming            XlsxReader (pull, IAsyncEnumerable) / XlsxWriter (forward-only, UTF-8)
L0  Primitives + OPC     CellValue, CellRef/RangeRef, ExcelSerialDate, FormulaText, OpcPackage
```

`Cellwright.Xls` is a second add-on over L2, pay-for-play like the others: the core does not pull
in the MS-CFB container, BIFF, or cryptography. It feeds the same `ISheetContentSink` as the
xlsx reader, so both formats land in the same model with a shared `WorksheetContentSink`. It also
has its own L1: `XlsReader` yields rows as `IAsyncEnumerable<RowData>`, just like `XlsxReader` —
the difference is that the workbook stream must be held in memory whole. The same package handles
password-protected packages: Excel wraps an entire xlsx in the same MS-CFB container, so the CFB
parser and crypto are already there while the core still pays nothing for them. For the same reason
VBA macros are extracted there: the VBA project is a compound file in both `.xlsm` and `.xls`.

Binary `.xlsb` lives there too, and not for convenience: it needs both halves that appear nowhere
else together. Its container matches xlsx — the same `OpcPackage` — while part records and formula
tokens match `.xls`. The Ptg decoder for both binary formats is literally one class: the token
dispatcher is shared, and differences — reference width, lengths of a few tokens, and the meaning
of byte `0x18` — are split by version. Details — [xls.md](xls.md), [xlsb.md](xlsb.md),
[vba.md](vba.md).

`Cellwright.Pdf` is a third add-on over L2 and the only one that pulls an external library: it
draws through Inkwright. The core knows nothing about it, so read/write-only users pay nothing.
From the core it needs exactly two internal pieces — the number-format renderer that already
displays values the way Excel does, and the workbook theme that resolves a theme slot to ARGB.
Everything else is the public model: `PageSetup`, styles, merges, pictures. Details —
[pdf-export.md](pdf-export.md).

`Cellwright.Formulas` is an add-on over L2, not part of the core: the core does not reference the
engine, so read/write-only users pay nothing. The engine attaches to the workbook through
`ConditionalWeakTable`, and the core gives it exactly three things: a cell-changed event, an
internal path to write a calculated value (without overwriting the formula), and sheet identity
so formulas bind to sheets by identity rather than name. Details — [formulas.md](formulas.md).

- **L0** — allocation-free value types: `CellValue` (24-byte tagged union), `CellRef`
  (`ISpanParsable`, `IUtf8SpanFormattable`), exact serial-date conversions (including the
  fictitious 29 Feb 1900 and the 1904 date system). The OPC layer is a thin wrapper over the
  async `ZipArchive` APIs in .NET 10 (`CreateAsync`, `OpenAsync`).
- **L1** — constant memory at any data volume. The writer serializes rows straight to UTF-8
  (custom `Utf8XmlWriter`: u8 literals, `IUtf8SpanFormattable`, `_xHHHH_` escaping); state is
  only the shared-string table. The reader is pull-based over async `XmlReader` with a reusable
  row buffer.
- **L2** — `Workbook` is a pure model: load reads the package whole and closes it; save runs the
  model through the L1 writer. No open handles between operations. Anything the model does not
  represent is kept as raw bytes and markup on load and written back on save — see
  [worksheet.md](worksheet.md).
- **L3** — `[ExcelRow]`/`[ExcelColumn]` attributes and a generator implementing `IRowSchema<T>`
  (static abstract members): headers, span-based `Read`/`Write`, column styles.

## Data-oriented storage (DOM)

No object per cell. A sheet stores three independent sparse “slices” (`SheetGrid<T>`): values
(`CellValue`), style indices (`int`), formulas (`string?`). Each slice is a column array; each
column is 512-row pages allocated lazily. Missing value = `default(T)`.

`Cell` and `CellRange` are readonly facade structs (sheet + reference pair) that write directly
into the slices. Mutators return the struct for chaining: `sheet.Cell("A1").SetValue(1).Bold()`.
The same pattern applies to `SheetRow` and `SheetColumn`: `sheet.Row(5).Hide()`,
`sheet.Columns(2, 8).Group()`.

## Preserving unrecognized markup

Copy-on-write at the package level: parts the model does not regenerate are copied byte for byte;
regenerated parts are rebuilt. A worksheet is special: it cannot be copied wholesale (cells are
the point of editing), so its unrecognized elements go back each into its own slot in the
`CT_Worksheet` schema. Worksheet parts that have a model but were not touched (comments with their
VML, tables) are also copied whole so yellow note boxes and table filter state survive. What is
preserved and what is deliberately dropped — [worksheet.md](worksheet.md).

## Styles

`CellStyle` and nested `FontStyle`/`FillStyle`/`BorderStyle`/`Alignment` are immutable records.
Record equality gives free interning: `StylePool` stores unique styles in a list; a cell holds
only an index. Indices 0–2 are fixed: default, date-time (builtin 22), duration (builtin 46) —
temporal values without an explicit style get these automatically.

Number-format classification (date/duration/number) on read uses builtin id and format-code parsing
(`NumberFormatClassifier`); the same logic turns serial numbers into `DateTime`/`TimeSpan` on read.

## Diagnostics instead of silent failure

Load is recovery-oriented: recoverable problems (broken references, unknown cell types, missing
shared-string indices, lost shared formulas) become `WorkbookWarning` with a code and cell
address — collected in `Workbook.LoadDiagnostics` and surfaced through `LoadOptions.OnWarning`.
`XlsxFormatException` is thrown only when the package is structurally broken.

Wrong-format files also get an explicit rejection instead of a deep parser error. By signature
and package contents the library recognizes: binary `.xls` and password-encrypted packages (both
MS-CFB containers — eight bytes cannot distinguish them, so the message names both and points to
`Cellwright.Xls`), `.xlsb` (same zip as xlsx — differs by the part the `officeDocument`
relationship targets), OpenDocument, `.docx`, `.pptx`, and plain zip without `[Content_Types].xml`.

## Foreign files as hostile input

Zip reports how large each entry expands to and is not obliged to tell the truth. A crafted package
can turn megabytes on disk into unbounded memory. `LoadOptions.Limits` caps a single part (512 MiB
by default) and their sum (2 GiB) — numbers far above any real workbook, so normal use never hits
them while a forgery gets a message instead of `OutOfMemoryException`. `PackageLimits.None`
removes the caps for a file you trust.

The per-part limit is enforced on actually unpacked bytes even for streaming parts: the zip
`Length` field cannot be trusted. The total budget counts parts that stay in memory: styles, the
string table, comments, and anything carried as raw bytes. Worksheets stream through the reader and
are never buffered, so a workbook with millions of rows does not count toward the aggregate budget.

`Cellwright.Xls` applies the same limits through `XlsLoadOptions.Limits` to the whole binary file
and each CFB stream. FAT, DIFAT, and miniFAT are validated before allocation; a loop or chain
shorter than the declared size is treated as format corruption.

The same rule governs macro decompression: six compressed bytes can expand to four thousand, with
its own limit (see [vba.md](vba.md)). XML is parsed with DTD disabled, so external entities cannot
exist by construction.

Long recalculation can be cancelled through `CalculateOptions.CancellationToken` — the engine
checks it between cells.

## .NET 10 / C# 14 features in use

| Feature | Where |
| --- | --- |
| Async `ZipArchive` API | `OpcPackage`, all package I/O |
| Extension members (C# 14) | Fluent styling `CellStyleFluent`, `CellFluent` |
| `params ReadOnlySpan<T>` | `XlsxSheetWriter.WriteRow(params ReadOnlySpan<CellValue>)` |
| u8 literals + `IUtf8SpanFormattable` | `Utf8XmlWriter`, cell serialization |
| `SearchValues` | XML escaping, sheet name validation |
| `FrozenDictionary` (+ alternate lookup) | Builtin formats, parsing `#DIV/0!` errors |
| `CollectionsMarshal.GetValueRefOrAddDefault` | Shared-string and style deduplication |
| Static abstract interface members | `IRowSchema<T>` |
| Incremental source generators | `Cellwright.Mapping.Generator` |
| `IsAotCompatible` + nullable everywhere | Both runtime packages |

## Deliberate v1 limitations

- The core does not calculate formulas: it stores and round-trips them as text. Calculation comes
  from `Cellwright.Formulas` (see [formulas.md](formulas.md)).
- Shared-formula members without the engine attached keep their calculated value but lose formula
  text and emit a warning; `Calculate()` expands groups into ordinary formulas.
- Insert/delete rows and columns and sheet rename rewrite formulas as text, not through parsing:
  the core does not depend on the formula package, so `FormulaText` in primitives recognizes
  references in text and rewrites only those, leaving everything else verbatim.
- AutoFit for widths and heights needs font measurement, so it lives in `Cellwright.Pdf` alongside
  metrics: `sheet.AutoFit()`, see [pdf-export.md](pdf-export.md). Without that package widths are
  set manually only.
- Theme colors are stored as theme slots, not resolved to pixels; `indexed` colors resolve
  through a fixed palette on read.
- One active sheet for the streaming writer (zip is strictly forward-only).
- Built-in charts are read (`sheet.Charts`), created (`sheet.AddChart`), and printed to PDF; raw
  `chartSpace` remains the source of truth and an untouched chart is written back byte for byte.
  A loaded chart cannot be edited — only deleted and recreated; chart sheets and pivot tables
  survive load-save without API.
- Password-encrypted packages are read and written through `Cellwright.Xls` (`SaveEncryptedAsync`);
  the core still knows nothing about encryption. All five schemes are read; one is written — agile
  AES-256/SHA-512, the one Excel writes.
- Macros are read (through `Cellwright.Xls`) but not written: the project part is copied byte for
  byte anyway, so there is nothing to regenerate — see [vba.md](vba.md).
- Sheet copy (`Sheets.Copy`) takes only what the model represents: drawings and other raw markup
  from the source worksheet part stay with the original.
- Differential formats (`dxf`) from a loaded file are parsed into a readable form (`rule.Style`),
  but the format table is written from the original markup: an untouched rule keeps its index; a
  rewritten style appends its format after preserved ones.
- The writer always emits a transitional package. Strict files are read, but unrecognized parts are
  lost on save — you cannot mix two conformance classes in one package. The full degradation list
  — [worksheet.md](worksheet.md).
- Binary `.xls` has no unrecognized-markup preservation: BIFF records are not self-contained and
  have cross-record offsets, so charts and pivots do not survive through it. Limits —
  [xls.md](xls.md).
- `.xlsb` is read but not written: save a read book as xlsx or `.xls`. Limits — [xlsb.md](xlsb.md).
- Pictures are read from all three formats and written only to xlsx. Charts and shapes beside them
  are not modeled: they travel with the untouched drawing part and are lost if worksheet pictures
  are changed. Details — [pictures.md](pictures.md).
- Where Excel diverges from ISO/IEC 29500, the library follows Excel, not the standard: the file
  was written by Excel and is opened by someone who expects Excel behavior. Divergences and the
  normative documents everything was checked against — [conformance.md](conformance.md).
