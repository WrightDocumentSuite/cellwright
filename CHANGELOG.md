# Changelog

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[semantic versioning](https://semver.org/).

While the major version is zero, breaking changes may land in minor releases: the public surface
is not frozen yet.

## [0.1.2] — 2026-08-01

The first release cut from the public repository. It touches no code beyond the PDF package's
dependency.

### Changed

- **The licence is Apache-2.0** instead of MIT: an explicit patent grant and notice requirements,
  neither of which MIT has. `Authors` is now Pavel Deecalov.
- **`Cellwright.Pdf` depends on `Inkwright` as a package** (pinned to `0.1.2`) rather than on
  sources from a neighbouring private repository. Building it publicly used to be impossible for
  anyone but the author, and the error message sent people to clone a repository they could not
  reach.
- **All five packages ship together.** Nothing blocks `Cellwright.Pdf` any more.
- The sources moved to <https://github.com/WrightDocumentSuite/cellwright>; CI no longer fetches
  anything from outside the repository.

### Unlisted

- `Cellwright` 0.1.0 and 0.1.1 — released under MIT and naming the old repository.

## [0.1.1] — 2026-08-01

Package metadata only; the code did not change.

### Changed

- The readme inside the package is now a separate English document with absolute links (the
  nuget.org page resolves no relative ones).
- The repository address in the metadata is <https://github.com/WrightDocumentSuite/cellwright>.

## [0.1.0] — 2026-08-01

The first release. The core — the `Cellwright` package — went to nuget.org; the other four built
and were verified but waited their turn (`Cellwright.Pdf` for `Inkwright` to be published, see
[docs/releasing.md](docs/releasing.md)).

### Added

- **Threaded comments** (`Cellwright`): `ThreadedComment`, `Person`, `sheet.ThreadedComments`,
  `book.People`. Threads are read, created and saved rather than merely surviving a round trip.
  The stub notes older readers rely on are regenerated from the threads on save — so editing
  ordinary comments no longer detaches threads from their cells, and the `ThreadedComments`
  warning has been removed for want of anything left to warn about.
- **Charts** (`Cellwright` + `Cellwright.Pdf`): read (`sheet.Charts` — kind, title, series with
  their references and cached points; the raw `chartSpace` is the source of truth and an untouched
  chart is written back byte for byte), created
  (`sheet.AddChart(ChartKind.Column, "D2:K17", "Data!$B$2:$B$10", …)` — column, bar, line, pie,
  doughnut, area; caches filled from the live cells, and the output passes the official OOXML
  validator) and printed to PDF (axes on round numbers, category labels, a legend, colours from
  the theme accents; a chart across a page break is reported as `chart-split` rather than drawn in
  half). Editing a loaded chart is still impossible — only reading, creating and removing; chart
  sheets round-trip without an API. Regenerating a sheet's drawing no longer loses its charts:
  they are written from the model alongside the pictures.
- **Dynamic array spill** (`Cellwright.Formulas`): a formula returning an array occupies its
  rectangle; an occupied cell, a foreign spill or a merge in the way yields `#SPILL!` in full.
  `A1#` references, `Cell.SpillRange`/`Cell.IsSpilled`, and a file written as an array formula
  with `XLDAPR` metadata — Excel opens it as dynamic, and a reload does not turn the spill into a
  CSE relic. Writing into a spilled cell breaks the spill, as it does in Excel.
- **Writing pictures to `.xls`** (`Cellwright.Xls`): images go into the BSE store, deduplicated by
  content, and the shapes are picture frames anchored in fractions of a cell. PNG, JPEG, BMP (as a
  DIB), EMF and WMF; GIF and TIFF are reported as `PictureFormatNotWritten`. Notes and pictures
  are written as one sheet drawing sharing a single block of shape identifiers.
- **Writing conditional formatting to `.xls`** (`Cellwright.Xls`): the classic `CondFmt`/`CF`
  records with a `DXFN` block. Comparisons and formula rules are written; the text rules and the
  ones about blank and error cells are lowered to the formula they are equal to — the very formula
  Excel itself puts into those rules in xlsx. Colour scales, data bars, icon sets, ranks, averages,
  duplicates and date periods live in `CF12` records the writer does not produce and are reported
  as `ConditionalFormatsNotWritten`; the sheet's other rules are still written.
- **Conditional formatting when printing to PDF** (`Cellwright.Pdf`): rules are worked out rather
  than ignored. Comparisons against literals, text rules, blanks and errors, duplicates and
  uniques, ranks, averages with deviations, all ten date periods, colour scales, data bars and
  icon sets need no formula engine. Expression rules do, and it is handed over through
  `PdfExportOptions.FormulaEvaluator`; without it such a rule is reported
  (`conditional-format-formula`) rather than treated as false. The blanket `conditional-format`
  refusal has been replaced by narrow codes.
- **Rich text in a cell** (`Cellwright`): `RichText`, `RichTextRun`, `cell.RichText` and
  `cell.SetRichText(...)`. The runs of a loaded file are now read and created from code, not only
  carried through a round trip. A run with no font of its own follows the cell's font; replacing
  the value drops the formatting along with the text it described.
- **AutoFit** (`Cellwright.Pdf`): `sheet.AutoFit()`, `AutoFitColumns`, `AutoFitRows`,
  `column.AutoFit()`, `book.AutoFit()` and `AutoFitOptions` with limits, row sampling and
  measuring settings. Fitting measures with what printing measures, so a fitted column prints
  without cutting its text. Wrapped cells, merged cells and hidden rows take no part in fitting a
  width — as in Excel.

### What each package holds

Below is the make-up of the first release rather than a list of changes: there is nothing to
compare it against.

#### Cellwright

The core: primitives, streaming, DOM.

- `CellValue` — a 24-byte tagged union with no boxing; dates at tick precision, serial dates of
  both epochs with the 1900 bug accounted for.
- Constant-memory streaming: `XlsxWriter` (forward-only, UTF-8 straight to the wire) and
  `XlsxReader` (`IAsyncEnumerable<RowData>`).
- A `Workbook`/`Worksheet` DOM with column slices in 512-row pages — not one object per cell.
- A complete worksheet: hidden rows and columns, grouping, auto-filter, hyperlinks, notes, Excel
  tables, data validation, conditional formatting, page setup, sheet and workbook protection, tab
  colour, zoom, document properties.
- Structural editing: inserting and deleting rows and columns, moving everything anchored to them
  and retargeting the whole workbook's formulas; reordering and copying sheets.
- Styles: immutable records with interning, fluent builders on extension members, all of §18.8 —
  pattern and gradient fills, diagonals, text rotation, theme colours, cell protection.
- Lossless editing of foreign files: charts, pivot tables, pictures, themes and `.xlsm` macros
  travel into the saved file as they were.
- Markdown export: `book.ToMarkdown()` and `ExportMarkdownAsync`.
- Diagnostics instead of silent failure: `WorkbookWarning` with codes, `LoadDiagnostics`,
  recovery-oriented loading; `LoadOptions.Limits` against zip bombs.

#### Cellwright.Formulas

- 447 functions, 352 of the 356 in §18.17.7 of ISO/IEC 29500 (all but four Japanese and Thai ones).
- Incremental recalculation over a dependency graph, cycle detection, shared-formula expansion.
- Defined names, structured references `Table[Column]`, 3-D ranges `Q1:Q4!A1`, references into
  closed workbooks from the `externalLink` cache, `LET` and `TEXT`.

#### Cellwright.Mapping

- `[ExcelRow]` + an incremental source generator → `IRowSchema<T>` on static abstract members.
- Zero reflection, Native AOT.

#### Cellwright.Xls

- Reading `.xls` (BIFF8, BIFF5, BIFF2–4) and `.xlsb` (BIFF12) into the same `Workbook`.
- Writing BIFF8: values, styles, notes through the Office Art layer, data validation, page setup,
  hyperlinks, Ptg formula encoding.
- Decrypting XOR/RC4/RC4 CryptoAPI and both MS-OFFCRYPTO schemes; writing an encrypted package
  through `SaveEncryptedAsync` under agile AES-256/SHA-512.
- Extracting VBA sources from `.xlsm` and `.xls` (MS-OVBA decompression).

#### Known limitations

Every deviation from the specification is declared in the documentation rather than passed over in
silence. The main ones as of this release:

- There is no `.xlsb` writer — save to xlsx or `.xls`.
- Conditional formatting is written to `.xls` except for the kinds that live in `CF12` records
  (colour scales, data bars, icon sets, ranks, averages, duplicates, date periods).
- A bare range reference (`=A1:A3`) collapses by implicit intersection rather than spilling; spill
  rectangles are known only after `Calculate()`.
- Pivot tables survive load and save but have no API; a loaded chart can be read and removed, but
  not edited.
- Phonetic hints (`rPh`, `phoneticPr`) survive a round trip but are not modelled, and are not
  restored once a cell's runs are rewritten.

The full lists are in [docs/](docs/).

[0.1.2]: https://github.com/WrightDocumentSuite/cellwright/releases/tag/v0.1.2
[0.1.1]: https://github.com/WrightDocumentSuite/cellwright/releases/tag/v0.1.1
[0.1.0]: https://github.com/WrightDocumentSuite/cellwright/releases/tag/v0.1.0
