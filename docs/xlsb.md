# Binary .xlsb (BIFF12)

The `Cellwright.Xls` package reads `.xlsb` — Excel’s binary format that appeared alongside xlsx in
2007. The loaded book is a normal `Workbook`: edit it with the same API and save to xlsx in one
line, so conversion comes out of the box.

Reading is read-only. Writing `.xlsb` is not supported: a loaded book can be saved as xlsx or `.xls`.

## Quick start

```csharp
using Cellwright;
using Cellwright.Xls;

Workbook book = await Spreadsheet.LoadAsync("report.xlsb");
book.Sheets[0]["B7"] = 42;

await book.SaveAsync("report.xlsx");
```

`Spreadsheet.LoadAsync` detects format from content, not extension. The zip container is the same for
xlsx and xlsb, so one signature is not enough: the fork follows the part name pointed to by the
`officeDocument` relationship. A part with extension `.bin` is the binary book, `.xml` is the usual
one.

The core without this package does not fail silently on `.xlsb`: `Workbook.LoadAsync` sees `.bin` in
the relationship and throws `XlsxFormatException` telling you to open the file through
`Spreadsheet.LoadAsync`.

## Encrypted workbooks

```csharp
Workbook book = await Spreadsheet.LoadAsync("locked.xlsb", new XlsLoadOptions
{
    Password = "secret",
});
```

An encrypted `.xlsb` uses the same MS-CFB container as an encrypted xlsx: Excel does not encrypt the
zip; it puts the whole package inside a compound file next to a stream describing the encryption.
Both schemes are supported — standard and agile — to the same extent as for xlsx
(see [xls.md](xls.md#encrypted-xlsx)). Saving the book back under a password is only possible as
xlsx: `SaveEncryptedAsync` encrypts the package, and there is no way to assemble the package in
binary form.

## What is read

| Area | Details |
| --- | --- |
| Values | All twelve cell records: blank, RK, double, inline string, shared string, rich text (with per-run `StrRun` formatting), boolean, error, and four formula kinds |
| Strings | `XLWideString` (UTF-16 with four-byte length) and the shared string table (`BrtSSTItem` with rich-text formatting runs) |
| Numbers | RK encoding same as BIFF8 — `fx100`/`fInt` bits |
| Dates | Both epochs, 1900 and 1904, from bit `f1904`; value type inferred from number format |
| Styles | Number formats, fonts, fills, borders, and the `cellXfs` table — into the shared `StylePool` |
| Formulas | Text decoded from Ptg tokens; cached value comes in the cell record itself |
| Shared and array formulas | `BrtShrFmla` and `BrtArrFmla`: group members take the master expression and expand it from their own position |
| Names | Defined names for book and sheet, with their expressions |
| 3-D references | Link table `BrtSupSelf`/`BrtExternSheet` with `Xti` structures |
| Sheet structure | Rows (height, hidden, outline, style), columns, merged cells, hyperlinks |
| Data validation | `BrtDVal`: ranges, type, operator, both condition formulas, input and error messages |
| Notes | Separate part by relationship: author list and text of each note |
| Conditional formatting | `BrtBeginCFRule`: value comparison, expression, text, blanks, errors, duplicates, average, top/bottom; graded activity formulas; color scales, data bars (including length limits), and icon sets (including `fGTE` thresholds) from their own records (`BrtBeginColorScale`/`BrtBeginDatabar`/`BrtBeginIconSet` + formula-aware `BrtCFVO`/`BrtColor`). `BrtDXF` properties are decoded into `DifferentialStyle` (font, fill, borders, and number format), attached by `dxfId`, and retained when converting `.xlsb` to `.xlsx`. |
| Tables | Table parts by relationship: name, range, header, totals, and column names |
| Structured references | `PtgList` rendered as `Table[Column]`, including `[#Data]`, `[#Headers]`, `[@]`, and column ranges |
| Pictures | Drawing part by relationship — the same DrawingML as xlsx; see [pictures.md](pictures.md) |
| Theme colors | `xl/theme/theme1.xml` is read; `BrtColor` with palette and theme slot is resolved |

## How BIFF12 differs from BIFF8

The two binary Excel formats are similar enough that half the code is shared, and they differ
exactly where it is costly not to notice.

**Record framing.** The id is one or two bytes, length one to four, both with seven significant bits
per byte (MS-XLSB 2.1.4). There are no `Continue` records at all: the body is always whole, which
makes reading simpler, not harder. Numbers below 128 are reserved for records the sheet is mostly
made of — a cell costs one header byte instead of two.

**A cell does not know its row.** Every cell record starts with a `Cell` structure (MS-XLSB 2.5.10):
column four bytes and format index 24 bits. The row is whichever one the preceding `BrtRowHdr`
declared. The sheet parser keeps a “current row”; no other reader in this library holds that state.

**References are wider.** `RgceLoc` is four bytes of row and two of column; `RgceArea` twelve
accordingly. Relative forms inside a shared formula wrap modulo the grid, as in `.xls`, but the
modulus differs: 1,048,576 rows and 16,384 columns instead of 65,536 and 256.

**Byte 0x18 means different things.** In BIFF8 it is `PtgSxName`; in BIFF12 it is `PtgList`, a
structured reference like `Table[Column]`. The tokens differ in length, so the fork by version is
mandatory: getting the size wrong misreads every token after it, not just that one.

**String in a formula is counted differently.** `PtgStr` in BIFF8 is a length byte and a flag byte
saying one-byte or two-byte characters (MS-XLS 2.5.240); in BIFF12 it is two length bytes and no
flag, because characters are always two bytes (MS-XLSB 2.5.98.74). The same trap as 0x18, and it
cost more: while strings were read the old way, the cursor landed mid-token, and every formula with
a text constant lost not the string but its entire text — 1,700 formulas in the corpus. The error
was invisible to the value oracle because the value comes from the record cache and does not depend
on whether the expression parsed.

**Colors are stored as values.** ARGB directly in the record instead of an index into a 56-color
palette — perhaps the one place the new format is simpler than the old.

## Recovery after errors

Loading is recovery-oriented, as elsewhere: recoverable issues become `WorkbookWarning` in
`Workbook.LoadDiagnostics`. The codes are the same as for other readers:

| Code | When |
| --- | --- |
| `EncryptedWorkbook` | File was encrypted; decrypted for reading |
| `SkippedSheet` | Chart, macro, or dialog sheet — this package does not model them |
| `UndecodedFormula` | Formula contains something that cannot be rendered; value kept |
| `UndecodedValidation` | Rule with a formula that could not be rendered; rule dropped |
| `CfRule12Skipped` | Conditional formatting rule of a kind the model does not know |
| `CommentTextMissing` | Note has no text record; left empty |
| `TextTooLong` | Text longer than 32,767 characters truncated |

Across the whole corpus `UndecodedFormula` does not occur once now: 11,696 formulas decode fully.

## Deliberate limitations

- **No writing.** The format is read but not written; save as xlsx or `.xls`.
- **No preservation of “unknown”**, as in `.xls`: records are not self-contained. Charts, pivot
  tables, and everything in the drawing part except pictures are skipped on read.
- **Chart, macro, and dialog sheets are skipped** with a warning. Sheet indices stay consistent —
  names and 3-D references still point where they should.
- **A gradient rule without its own records** — a color scale with no following `BrtCFVO` — is
  reported as `CfRule12Skipped`, not turned into a rule without a scale.
- **Cross-workbook table references are not rendered.** A `PtgList` token with the `nonresident`
  flag points outside the file; the text of such a formula is not reconstructed.

## How it is verified

- **Corpus**: 66 `.xlsb` files from ExcelDataReader test data and neighboring libraries — both date
  epochs, rich text, column widths, collapsed rows, double precision, broken OADate, a 10×10,000
  sheet. 53 read; of the rest 12 are encrypted and one is an intentionally empty zip.
- **Value oracle**: the same files are read by ExcelDataReader and values compared cell by cell —
  296,333 comparisons, zero mismatches.
- **Completeness check**: separately asserts that every sheet the oracle sees exists in the loaded
  book. Value comparison matches sheets by name and silently skips missing ones, so a book that
  lost all sheets would pass with zero mismatches — exactly how a `BrtBundleSh` parsing bug looked
  before this check found it.
- **Formula oracle**: there is no external decoder of formula text for this format — the library
  serving as the value oracle does not parse formulas at all. So verification is indirect: decode,
  recalculate with the engine, compare to the cached value from the file. 11,339 formulas, full
  match. The check is weaker than a direct one — two different expressions can yield the same
  result — but a misread token stream is caught immediately: a reference shifted by one row does not
  land on the same number.
- **Cross-format oracle**: 64 corpus books also exist as `.xls` or `.xlsx`, the only way to read the
  same book twice by independent paths. Compared are sheet names, values (296 thousand cells),
  **formula texts** (11,694), merges, hyperlinks, notes, data validation rules, conditional formats,
  and tables. This comparison found everything fixed in this phase: missing data validation on 13
  sheets, a lost note, and a divergent `PtgStr`. Pairs with different content under the same name
  (four of them) are filtered by sheet list mismatch and counted separately.
- **Encrypted workbooks**: ExcelDataReader matrix — standard scheme with 128, 192, and 256-bit keys
  and agile with all hashes; each checks content, not just absence of exception. Files on DES, 3DES,
  and RC2 are rejected with the algorithm name, as for xlsx.
- **Golden tests** on named fixtures: tab order, both epochs, rich text, three error kinds, a formula
  with its text and result, number precision at range edges, merges, hidden rows and columns, a
  ten-thousand-row sheet.

### Where coverage is synthetic only

This deserves to be stated plainly because a green run does not show it. **There is no conditional
formatting or tables in any of the 66 corpus files** — nor hyperlinks, for that matter.
Those readers — including color scales, data bars, and icon sets — are verified by fixtures built
byte-by-byte in tests from MS-XLSB, and that only shows the code agrees with the spec reading
baked into the same test. The `BrtDXF` property bytes were additionally checked against files
created by Excel 16, but the committed regression is still synthetic; neither check replaces a
real conditional-formatting file in the corpus.

By contrast: values, formulas, notes, and data validation are touched by the corpus, and the
cross-format oracle works there — independent reading of the same content by different code.
The gap between these two kinds of confidence is substantial, and they should not be relied on
equally.
