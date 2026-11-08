# Legacy .xls (BIFF)

The `Cellwright.Xls` package reads and writes Excel's binary format — the `.xls` that lived from
1987 through 2007. A loaded workbook is a regular `Workbook`: edit it with the same API, save to
xlsx in one line, so conversion works out of the box.

The package is separate by design (pay-for-play): the core does not pull in the CFB container,
BIFF, or cryptography. It references `Cellwright` (reuses `ISheetContentSink`, `StylePool`,
`SheetGrid` via `InternalsVisibleTo`) and `Cellwright.Formulas` (the formula text parser is needed
to encode formulas back into Ptg on write).

## Quick start

```csharp
using Cellwright;
using Cellwright.Xls;

// Reading — returns a regular Workbook
Workbook book = await XlsWorkbook.LoadAsync("report-2003.xls");
book.Sheets[0]["B7"] = 42;

// Convert to the modern format
await book.SaveAsync("report-2026.xlsx");

// Or back to binary
await book.SaveXlsAsync("report-2003-updated.xls");
```

If the file format is unknown in advance — and you cannot trust the extension; files named `.xls`
with xlsx contents are common — `Spreadsheet.LoadAsync` detects it by signature:

```csharp
Workbook book = await Spreadsheet.LoadAsync(uploadedFile);   // xlsx or xls, either way
```

The binary file, CFB directory and streams, BIFF cells and strings, pictures, and embedded objects
use the same `DocumentLoadBudget` as xlsx:

```csharp
var options = new XlsLoadOptions
{
    Budget = DocumentLoadBudget.Default with
    {
        MaxInputBytes = 64 * 1024 * 1024,
        MaxPartBytes = 64 * 1024 * 1024,
        MaxInflatedBytes = 256 * 1024 * 1024,
        MaxCells = 2_000_000,
    },
};
Workbook book = await Spreadsheet.LoadAsync(uploadedFile, options);
```

The check runs before allocating an array from a declared size. `PackageLimits` and `Limits` remain
compatible; `PackageLimits.None` disables every ceiling as before. Prefer a finite custom budget
for uploads. See [Loading untrusted spreadsheets](loading-untrusted-input.md).

Without this package, the core does not fail silently on a binary file: `Workbook.LoadAsync`
recognizes the CFB signature and throws `XlsxFormatException` with text saying to add
`Cellwright.Xls`. The same container also carries password-protected xlsx (see
[Encrypted .xlsx](#encrypted-xlsx)), so the message mentions both possibilities.

## Encrypted workbooks

```csharp
Workbook book = await XlsWorkbook.LoadAsync("protected.xls", new XlsLoadOptions
{
    Password = "secret",
});
```

All three "password to open" schemes are supported: XOR obfuscation (Excel 95 and earlier), 40-bit
RC4 via MD5 (Excel 97), and RC4 CryptoAPI via SHA-1 (Excel 2000+). The default password
`VelvetSweatshop` is tried first, so files that Excel encrypts "silently" open without any
configuration. Writing is always unencrypted; `EncryptedWorkbook` warns on load.

All three schemes were verified line-by-line against MS-OFFCRYPTO (sections 2.3.5–2.3.7). This is
the code where a mistake does not throw: a wrong key decrypts the file to garbage instead of
refusing, so verification against the spec text matters more than a corpus run.

### Encrypted .xlsx

A package protected with a password to open is also this package's concern, even though inside it
is modern xlsx. Excel does not encrypt the zip: it puts the whole package inside the same MS-CFB
container, alongside a stream describing the encryption. So such a file starts with exactly the same
eight bytes as a 2003 workbook, and anything that looks only at the signature treats it as legacy.

```csharp
Workbook book = await Spreadsheet.LoadAsync("locked.xlsx", new XlsLoadOptions
{
    Password = "secret",
});
```

Both schemes: standard Office 2007 (AES-ECB, key from 50,000 SHA-1 rounds) and agile Office 2010
and later (algorithms declared in an XML descriptor in the file; the package is decrypted in
4096-byte segments, each with its own vector). Ciphers AES, DES, 3DES, and RC2 in CBC and CFB
modes; hash SHA-1, SHA-256, SHA-384, SHA-512, or MD5. Key length is validated against what the
named cipher accepts: a mismatch means the descriptor and data disagree, and decryption would yield
plausible garbage. The extensible scheme with an external crypto module and DESX cipher are
rejected with the cipher name in the message.

The agile descriptor is untrusted XML, not just algorithm configuration. In addition to the shared
XML budget, it has a 1 MiB format ceiling checked from the CFB directory before its stream or any
XML attribute is materialized. Every Base64 salt, verifier, wrapped-key, and HMAC attribute is also
capped at 4 Ki encoded characters before allocating its decoded buffer. Real descriptors use only a
small fraction of either limit; an oversized descriptor or field is rejected as `XlsFormatException`.

Agile files have an HMAC over the encrypted package (MS-OFFCRYPTO 2.3.4.14). It is verified: a
file modified after writing still reads, but comes with a `PackageIntegrityFailed` warning.

The decrypted package enters the same core workbook loader, so preservation of unknown parts works
as usual. The outer CFB directory/streams, encryption descriptor, decrypted OPC parts and any
nested resources share one `DocumentLoadBudget` state; format detection does not reset its counters.
Password stretching and package decryption observe the caller's cancellation token. Loading warns
with `EncryptedWorkbook`.

`XlsWorkbook.LoadAsync` does not guess on such a file: it says it is an encrypted package and
points you to `Spreadsheet.LoadAsync`. Without this package, the core is honest about the
container — legacy workbook or encrypted package; the eight bytes cannot tell them apart.

### Writing with a password

```csharp
await book.SaveEncryptedAsync("locked.xlsx", "secret");
```

It writes **one** scheme out of many the format allows: agile, AES-256 in CBC with SHA-512, 100,000
rounds, 16-byte salts, integrity HMAC — exactly what Excel writes today. Reading must handle
whatever is in the file; there is no obligation to write weaker ciphers, and choosing weaker ones
would be a way to get it wrong.

Inside is a regular xlsx, built by the same writer as always: nothing extra for the password,
unknown parts (charts, pivots, macros) pass through. An empty password is rejected with
`ArgumentException` — a file that opens without a prompt is not a loud encryption failure, but a
quiet one.

There is no streaming here and cannot be: the integrity value is computed over the finished
ciphertext, so the package is assembled in memory in full.

Password-protected `.xls` writing is not supported: BIFF encrypts record bodies selectively, and
that would touch every record `BiffWriter` writes.

## What is read

| Area | Details |
| --- | --- |
| Container | MS-CFB v3/v4: FAT, DIFAT, miniFAT, directory; streams `Workbook` (BIFF8) and `Book` (BIFF5) |
| Values | RK, MulRk, Number, BoolErr, Blank, MulBlank, LabelSst, Label, RString (with per-run rich text formatting) |
| Strings | SST with `Continue` stitching (including strings split mid-character and rich-text formatting runs) |
| Formulas | Rgce → text: all operators and operands, 3-D references, arrays, shared formulas via PtgExp |
| Link table | SupBook/ExternSheet/ExternName: references to another workbook with its path and sheet names, add-in function names (Analysis ToolPak's `EOMONTH` and the like) |
| Styles | Font, Format, XF, Style, Palette → `CellStyle` in the workbook's shared pool |
| Sheet | Row, ColInfo, MergeCells, Window2, Pane, Scl, HLink, HLinkTooltip, DV, SheetExt, Protect |
| Hyperlinks | OLE monikers per MS-OSHARED 2.3.7: URL, file, composite, anti-, item; target stored as string; tooltip from a separate record |
| Document properties | Streams `\005SummaryInformation` and `\005DocumentSummaryInformation` (MS-OLEPS): title, author, company, dates, custom fields |
| Nested substreams | Embedded chart — substream inside a sheet: skipped entirely; records after it (comments, window, merges) are read as part of the sheet |
| Comments | Note + Obj (ftCmo) + TxO: anchor, author, and text from Continue fragments → `Worksheet.Comments` |
| Print | Setup, margins, Header/Footer (+ HeaderFooter 0x089C), HCenter/VCenter, PrintGrid/PrintRowCol, WsBool, page breaks |
| Conditional formatting | CondFmt/CF (comparison and formula, DXFN → `DifferentialStyle`), CondFmt12/CF12 (activity and CFVO formulas, theme colors with tint, color scales, data bars including minimum/maximum lengths, icon sets including strict/inclusive thresholds) |
| Workbook | BoundSheet8 (names, three visibility states), Date1904, CodePage, Lbl, SupBook/ExternSheet |
| Encryption | FilePass: XOR, RC4, RC4 CryptoAPI; and encrypted xlsx package in the same container |
| Macros | VBA project from `_VBA_PROJECT_CUR` storage — separate API, not through the model ([vba.md](vba.md)) |

Dates and durations are recovered the same way as in xlsx: a number becomes `DateTime` or
`TimeSpan` when its number format is date or `[h]:mm`. Cached formula values go through the same
classification.

Built-in Excel names are applied to the model, not left as names: `_xlnm.Print_Area` becomes the
sheet's print area, `_xlnm._FilterDatabase` the auto-filter range (and sets `AutoFilterIsActive`,
which `SUBTOTAL` depends on).

Comments, conditional formatting rules, and print settings become model-owned, not file-owned. A
BIFF workbook has no saved markup to replay on save, and the xlsx writer generates these parts only
for what the model owns — otherwise conversion would silently drop them.

### Row streaming

When you only need values, `XlsReader` yields rows one at a time without building a model:

```csharp
await using XlsReader reader = await XlsReader.OpenAsync("export.xls");
await foreach (RowData row in reader.ReadRowsAsync(reader.Sheets[0].Index))
    Console.WriteLine(row.GetValueOrBlank(0));
```

`RowData` is a view over a reused buffer: copy the row (`ToArray`) if you need it after the next
enumeration step. Empty rows are skipped, so look at `RowNumber`.

Unlike xlsx streaming, the workbook stream is read into memory in full: the MS-CFB container is a
chain of sectors scattered through the file, and the sheet table at the start points to byte
offsets, so there is nothing to read "as it arrives." Streaming saves not file I/O but the model:
no styles, formulas, or per-cell objects.

Cell order within a row block is not fixed by the format — real files write by column and return
to an already-started row — so the sheet is scanned once "dry" to learn each row's last cell.
That pass reads only record headers and costs a small fraction of the main read, but a row is
yielded only when nothing more will be added to it.

### BIFF5 and BIFF2–4

BIFF5 (Excel 5.0/95) is read the same way: same container, stream named `Book`, single-byte strings
in the declared code page. BIFF2–4 is a file without a container at all — a bare record stream with
one sheet; values and basic number formats are read.

Formulas from those generations are not decoded: their encoding differs, and the cached result is
already in the file, so the cell keeps the last displayed value. Loading reports this with a
`LegacyBiff` warning.

### Code pages

Single-byte BIFF2–5 strings are decoded using the code page declared in the file. Most such pages
in .NET are available only after registering a provider — that is the caller's responsibility;
the library does not do it for you:

```csharp
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
```

Without registration, text is read as Windows-1252 and a `UnknownCodePage` warning appears.

## What is written

`SaveXlsAsync` produces BIFF8 (Excel 97-2003): global substream with styles, palette, SST, and
names, then sheets in 32-row blocks with `DBCell` index, all in a minimal CFB container.

Numbers are written as RK when four bytes hold them exactly, and as `Number` otherwise. Formulas
are encoded back to Ptg through the `Cellwright.Formulas` parser; what has no token in this format
is written as the last value with a warning.

| Area | Details |
| --- | --- |
| Values | RK/Number, LabelSst with shared string table (including rich-text formatting runs), BoolErr, Blank; dates as numbers with their format |
| Formulas | Ptg via `Cellwright.Formulas` parser, including 3-D references (SupBook + ExternSheet) |
| Styles | Font, Format, XF, Style, 56-slot palette with color quantization |
| Sheet | ColInfo, Row, MergeCells, Window2/Pane/Scl, HLink, Protect |
| Print | WsBool, Setup, margins, headers/footers (including even and first page), breaks |
| Data validation | DVal + DV: types, operators, messages, lists — both by reference and literal enumeration |
| Comments | Drawing layer: MsoDrawingGroup in workbook, MsoDrawing/Obj/TxO and Note on sheet |

In the binary format, a comment is not a record but a shape: a text box in the Office Art tree, an
object that names it, text that fills it, and only then a record tying it to a cell. The smallest
tree Excel accepts is written: one drawing per sheet, a group shape, and one text box per comment,
with the default box above-right of the cell. Geometry of original boxes is not preserved when
converting from xlsx.

## Error recovery

Loading is recovery-oriented, like the core: recoverable issues become `WorkbookWarning` in
`Workbook.LoadDiagnostics`, and `XlsFormatException` is thrown only when the file is hopelessly
broken. Codes specific to the binary format:

| Code | When |
| --- | --- |
| `EncryptedWorkbook` | File was encrypted and decrypted for read; `.xls` will be written back unencrypted |
| `LegacyBiff` | Workbook from Excel 4.0 and earlier: values read, formulas not |
| `MissingEof` | Substream has no end marker; next substream's start used as boundary |
| `BadSheetOffset` | Sheet offset in BoundSheet8 is wrong; sheet found nearby or read empty |
| `SkippedSheet` | Chart, macro, or dialog sheet — this package does not model them |
| `UndecodedFormula` | Formula has a construct with no renderer; value kept |
| `UnknownCodePage` | Declared code page unavailable; text read as Windows-1252 |
| `TruncatedSst` | String table promises more strings than it contains |
| `UnknownHyperlinkTarget` | Link described by an unparseable moniker; display text kept |
| `PackageIntegrityFailed` | Encrypted package does not match recorded HMAC: file modified after writing |
| `CommentTextMissing` | Note references a drawing object without text; comment left empty |
| `CfRule12Skipped` | Conditional formatting rule with no place in the model was not read |
| `GridTooLarge` | On write: content beyond 65536×256 does not fit the legacy grid |
| `TextTooLong` | Text longer than 32767 characters was truncated |
| `UnwritableFormula` | On write: formula or validation rule has no BIFF encoding |
| `ConditionalFormatsNotWritten` | On write: rule type with no classic record (scale, data bar, icons, rank, average, duplicates, period) |
| `ThreadedCommentsNotWritten` | On write: threaded comments have no representation in the legacy format; convert them to classic comments to retain their text |

## Known limitations

- **No preservation of "unknown" content.** On the xlsx side, unknown package parts are carried
  byte-for-byte; in BIFF that is impossible — records are not self-contained and have cross-offsets.
  Charts, pivot tables, VBA, and the rest of the Escher layer are skipped on read and not restored
  on write. Macro source can still be extracted, just not through the model:
  `VbaProject.LoadAsync` reads it directly from the file — see [vba.md](vba.md).
- **Pictures are read and written.** Writing puts images in the BSE store inside
  `MsoDrawingGroup` — identical bytes stored once no matter how many shapes reference them — and on
  the sheet draws a picture frame with `pib` and anchor in cell fractions (1024ths of column width,
  256ths of row height), the inverse of what the reader parses. Comments and pictures live in one
  sheet drawing and share one shape ID block — two drawings per sheet would send the file into
  recovery. Formats: PNG, JPEG, BMP (as DIB, without file header), EMF and WMF (uncompressed); GIF
  and TIFF have no place in the Escher layer and are reported with `PictureFormatNotWritten`. Image
  names are not carried over.
- **Chart, macro, and dialog sheets are skipped** with a warning: the model does not know them.
  Sheet indices stay consistent — names and 3-D references still point where they should. An
  embedded chart on a sheet is lost too, but only that: the rest of the sheet content is read on
  both sides of its substream.
- **Comments**: read and written for BIFF8; Note from older generations is structured differently
  and is not parsed. Read comments belong to the model, so saving to xlsx regenerates comment parts
  and VML — with default box geometry.
- **Threaded comments** were introduced after the `.xls` format and are not written. Each affected
  sheet raises `ThreadedCommentsNotWritten`; convert the messages to classic comments before export
  when their review text must survive.
- **Conditional formatting is partially written.** Classic `CondFmt`/`CF` records know exactly two
  condition kinds: cell comparison with one or two formulas, and a formula that is true or false on
  its own. Both are written, together with differential format (font, color, borders, number format)
  in the `DXFN` block.

  Text rules and rules about blank and error cells have no record of their own, but each is exactly
  equivalent to a formula — and that is what Excel itself puts in the `formula` attribute of the
  same rules in xlsx. So they are lowered to a formula rather than dropped: `contains "x"` becomes
  `NOT(ISERROR(SEARCH("x",A1)))` and reads back in any reader of classic records. The rule type
  becomes "expression" — notation is lost, not meaning.

  Color scales, data bars, icon sets, `Top 10`, "above average", duplicates, and date periods are
  not carried over: they live in `CF12` records, which the writer does not create. Each such rule
  gets `ConditionalFormatsNotWritten` — other rules on the sheet are still written.

  A `CondFmt` group counts its rules as a field allowing one to three, so a fourth rule over the
  same range goes in a separate group, not lost. Rules are grouped by matching range sets — there
  is no alternative: one record names one cell set for all rules beneath it.
- **CFEx records** (Excel 2007 extensions on classic groups) are not parsed: priority and "stop if
  true" for such rules come from order in the file.
- **BIFF5 and BIFF2–4**: values and basic formats only; formulas as cached value. Streaming is not
  available for them: without a sheet table there is nothing to seek to.
- **Ptg operand classes** on write are set to the value class. Excel rebuilds them on first
  recalculation, and every tested reader tolerates that, but a formula whose meaning depends on
  class difference may evaluate differently in Excel until the first recalc.
- **External workbook references are read but not evaluated.** Formula text names the workbook, its
  path, and sheet — `'[C:\Reports\LastYear.xls]Sheet1'!$D:$I` — so you can see where it points, but
  the engine does not leave the current workbook and returns `#NAME?`. A name defined in another
  workbook becomes `#REF!`: Excel caches exactly that in such a formula when the file is no longer
  there. Add-in functions are a separate case and resolve to their own name: Analysis ToolPak's
  `EOMONTH` is stored the same way but means a function, not a reference.
- **`.xls` is always written unencrypted**, even if the source was encrypted: BIFF encrypts record
  bodies selectively, and that would touch every record the writer emits. Password-protected xlsx, by
  contrast, is both read and written — `SaveEncryptedAsync`.

## How it is verified

- **Corpus**: 844 `.xls` files from ExcelDataReader, NPOI, EPPlus, and Aspose.Cells test data.
  All load or fail with a clear error; no mid-parser crashes.
- **Value oracle**: the same files are read by ExcelDataReader and values compared cell by cell —
  49.8M comparisons, 27 mismatches (0.00005%), and in each this package is more accurate: wrong
  values the oracle returns as blank, and array formula cache it does not reach.
- **Formula oracle**: decoded expression is evaluated by the engine and compared to cached result
  from the file — 98.2% match on 54k formulas. The remainder is concentrated in 21 of 293
  workbooks, two of which account for most of it; volatile functions and `#REF!` references are
  excluded from the sample — their cache is not reproducible by definition.
- **Formula text oracle**: the same corpus is decoded in parallel with NPOI and the two texts
  compared expression by expression — 99.1% match, differences in 3 of 119 files. Comparison with
  an external decoder catches what neither round-trip nor the value oracle sees: a "reader and
  writer, equally wrong" pair where the file disagrees with Excel but agrees with itself. That found
  swapped `PtgUplus`/`PtgUminus` (sign of each unary term flipped on both read of foreign files and
  write of our own) and a formula without a computed value that the writer declared string without
  adding a `String` record. Remaining differences are our defensive parentheses around unary minus
  and preserved absoluteness flags on whole-column references that NPOI drops.
- **Round-trip**: what this package writes is read by itself and, more importantly, by two external
  libraries — writer agreeing with its own reader only proves shared misconceptions. NPOI is
  stricter here: it parses every record encountered, does not skip unknown, and caught three
  malformed records in the global substream.
- **Streaming**: rows from `XlsReader` are compared to the model of the same files across the full
  corpus — the two paths to one file share nothing except the result.
- **Units**: RK codecs, `Continue` boundaries, all string forms, code pages, three encryption
  schemes. Pure RC4 has no file in the corpus and no available library writes it, so the fixture is
  built in tests from the spec.
- **Encrypted packages**: ExcelDataReader matrix — standard scheme with 128, 192, and 256-bit keys
  and agile with all five hashes and all four ciphers (AES, DES, 3DES, RC2); each checks cell
  content, not just absence of exception. Separately verified: HMAC matches on untouched file and
  fails if a ciphertext byte is flipped — integrity check that always stays silent checks nothing.
- **Password writing** — the case where round-trip proves almost nothing: reader and writer with a
  shared misconception agree perfectly, and the file is one nobody else can open. So what we write is
  opened by **NPOI** (via `POIFS.Crypt`, including rejection on wrong password) and
  **ExcelDataReader** — two external MS-OFFCRYPTO reads on different paths. Plus: our HMAC is
  verified by the same code that verifies HMAC on Excel-written files — if construction diverged
  from the spec, our file would fail a check that foreign files pass. Also tested: workbook longer
  than one segment (otherwise vector numbering is unused — first 4096 bytes decrypt on any error
  there), pivot preservation inside encryption, and two writes of the same workbook with the same
  password produce different bytes: identical would mean no randomness.
- **Hyperlink oracle**: link targets compared to NPOI across the corpus. Hyperlink structure is the
  only one in the format where all fields are optional and consecutive without separators, so a
  misread flag does not drop one field but shifts all following ones; shifted reading yields a
  plausible string, catchable only by an external reader. That found the "text supplied by author"
  flag was taken for "text present."
- **Document properties oracle**: title, author, company, etc. compared to NPOI — 2,199
  comparisons, no mismatches. That found searching the stream catalog by name picks properties of
  an *embedded* Word document, not the workbook itself.
- **Comments**: POI fixtures with known content (`SimpleWithComments.xls`, `DrawingAndComments.xls`)
  — authors, texts, anchors, and survival of xlsx conversion; comments written by this package are
  read by NPOI.
- **Conditional formatting**: POI fixtures with rules spelled out in its own tests — operators,
  formulas, bold from DXFN, color scales and icon sets with their thresholds.

## Performance

50,000 rows × 4 columns, both readers on the same bytes
(details in [benchmarks.md](benchmarks.md)):

| | Time (median) | Allocations |
| --- | ---: | ---: |
| **Cellwright: read to DOM** | **51 ms** | **17.7 MB** |
| **Cellwright: row streaming** | 128 ms | 16.6 MB |
| ExcelDataReader 3.7: read | 142 ms | 96.1 MB |
| **Cellwright: write** | 83 ms | 10.6 MB |

```bash
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks -- --filter *XlsBenchmarks*
```
