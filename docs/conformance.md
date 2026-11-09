# Specification conformance

The library is written against normative documents, not by observing files in the wild. This page
lists which documents apply, what has been verified line by line, and where Excel diverges from the
standard — because in those places the correct answer is Excel's, not the standard's.

## Reference documents

| Document | What we implement from it |
| --- | --- |
| ISO/IEC 29500-1 (ECMA-376 Part 1) | SpreadsheetML: workbook, sheet, style, and formula markup; external links and their cache (§18.14); functions §18.17.7 — 352 of 356 |
| ISO/IEC 29500-1, DrawingML | Pictures and anchors (§20.5), theme schema (§20.1.6.2), tint (§18.8.3) |
| ECMA-376 Part 2 | OPC: part names, content types, relationships |
| ISO/IEC 29500-3 | Markup Compatibility: `Ignorable`/`mc:AlternateContent` passthrough, `MustUnderstand` diagnostics; `Choice`/`Fallback` branch selection is not performed |
| MS-XLSX | Office extensions on top of 29500 (`x14`, `x15`) |
| MS-XLS | Binary `.xls`: BIFF8 records, Ptg tokens |
| MS-XLSB | Binary `.xlsb`: BIFF12 records, Ptg tokens, data validation, comments, conditional formats, tables |
| MS-OFFCRYPTO | Five read schemes (XOR, RC4, RC4 CryptoAPI, standard, agile); password-protected package write — agile |
| MS-CFB | Compound-file container: read v3/v4, write v3, FAT/DIFAT/miniFAT and chain-cycle protection |
| MS-ODRAW | Office Art layer in `.xls`: BLIP store, picture shapes, anchors |
| MS-OSHARED | `.xls` hyperlinks: OLE monikers |
| MS-OLEPS | `.xls` document properties: property set streams |
| MS-OVBA | VBA source extraction |
| GitHub Flavored Markdown Spec | Markdown export: tables and separator row (§4.10) |
| CommonMark (within GFM) | Backslash escaping (§6.1), links and images (§6.6, §6.7) |
| MS-OI29500 | Office deviations from the standard — section below |

Open Microsoft specification revisions are tracked: last review — MS-XLSB 21.1 (2025-11-13) and
MS-OFFCRYPTO 14.0 (2026-02-17). MS-XLSB major release 21.0 added only pivot-table records
(`BrtSxvdSubtotal*`, `BrtPivotRuleFilterSubtotal*`, `BrtPivotCacheDataSource`) — the reader skips
them like any unrecognized record number; MS-OFFCRYPTO 14.0 added OneNote sections. Neither affects
implemented areas.

## Where Excel diverges from the standard

MS-OI29500 documents 203 functions whose Office behavior differs from 29500. One hundred seven are
implemented here. Each deviation is checked in
[ConformanceTests.cs](../tests/Cellwright.Formulas.Tests/ConformanceTests.cs); of thirty tested
assertions, twenty-one already passed before the review and nine required fixes.

The library follows Excel everywhere, not the standard: the file was written by Excel and opened
by someone who expects Excel's answer. Where the two disagree, reading by the standard gives a
defensible result, but not the one in the cell.

| Function | Standard | Excel (and us) | Section |
| --- | --- | --- | --- |
| `CEILING`, `FLOOR` | Rounds toward zero when signs differ | `#NUM!` for a positive number with a negative step; a negative number with a positive step is calculated | 2.1.895, 2.1.954 |
| `LOG` | Does not specify base 1 | `#DIV/0!` | 2.1.986 |
| `POWER(0,0)` | `#DIV/0!` | `#NUM!` | 2.1.1018 |
| `PMT` when `nper=0` | Unspecified | `#DIV/0!` | 2.1.1016 |
| `XNPV` when `rate<=0` | Unspecified | `#NUM!` | 2.1.1071 |
| `INDIRECT` longer than 255 characters | Unspecified | `#VALUE!` | 2.1.974 |
| `PERCENTILE` of an empty array | Undefined | 0 | 2.1.1014 |
| `DAYS360`, end of February | Moves only the 31st | Moves end of February too | 2.1.923 |
| `ADDRESS` when row or column < 1 | `#NUM!` | `#VALUE!` | 2.1.877 |
| `MATCH` with no match | `#NUM!` | `#N/A` | 2.1.990 |
| `SUMPRODUCT` of mismatched shapes | `#NUM!` | `#VALUE!` | 2.1.1049 |
| `MID`, `FIND`, `SEARCH`, `REPLACE` | Zero-based positions | One-based | 2.1.992, 2.1.950, 2.1.1037, 2.1.1029 |
| `STDEV`, `VAR` with one value | Unspecified | `#DIV/0!` | 2.1.1042, 2.1.1062 |
| `HOUR` | Range 0–59 | Range 0–23 | 2.1.962 |
| `POISSON` with mean 0 | `#NUM!` | Calculates | 2.1.1017 |
| `CLEAN` | Removes U+0000–001F | Also removes U+0080–009F | 2.1.898 |
| Text longer than 32,767 characters | Unspecified | `#VALUE!` | 2.1.868 |
| `#GETTING_DATA` | Not an error value | Is an error value | 2.1.869 |

The remaining deviations among the 107 are editorial: the standard does not name the angle unit, or
says "count" where it means "sum", or requires an iterative method where Office may calculate
differently. Behavior does not depend on them.

## Where we deliberately depart from the letter of the standard

- **Packages are always written transitional.** Strict files are read, but on save unrecognized
  parts are dropped: mixing two conformance classes in one package is forbidden (29500-1 §2.2).
- **Part names are compared ASCII case-insensitively**, as 29500-2 §6.2.2.3 requires, not via the
  platform's `OrdinalIgnoreCase`. The difference is not academic: a shared folding rule maps Kelvin
  sign U+212A to K, and two different package parts become one.
- **`calcChain.xml` is ignored and not created**: our own dependency graph is built from formulas,
  not from Excel's calculation-order cache.
- **Ptg operand classes on `.xls` write** are set by value class; Excel rebuilds them on first
  recalculation.
- **Markdown escaping does not escape every punctuation character**, although §6.1 allows it —
  only what breaks structure: pipe, backslash, and line break. A source where every asterisk is
  backslash-escaped is unreadable, and Markdown is chosen for its source — see
  [markdown.md](markdown.md).

## How conformance is verified

The specification says what to expect; a corpus says what actually appears. Neither alone is
enough — files diverge from the spec, and the corpus does not cover corners that simply are not in
it.

- **Line-by-line verification** where a mistake does not throw: all cryptography, BIFF12 record
  framing, moniker structure. Such code "works" on the corpus and silently lies on a file outside it.
- **Corpus oracles**: cell values against ExcelDataReader (49.8 million comparisons for `.xls`,
  296 thousand for `.xlsb`), formula text against NPOI, hyperlinks against NPOI, document properties
  against NPOI (2,199 comparisons), images against NPOI (337 placements in `.xls`, 151 in xlsx),
  theme colors against NPOI (1,428 slots), our encrypted package — against NPOI and ExcelDataReader
  together.
- **Synthetic fixtures** where the corpus is silent: hyperlink flag combinations nobody saved,
  encryption descriptors with foreign algorithms, packages with part names at the edge of the rules.

## Encrypted-package mutation checks

The encrypted-write integrity test checks for `PackageIntegrityFailed` after a
ciphertext mutation. Fresh salts and keys make the downstream damage vary: a
mutation may break ZIP framing or XML inside an otherwise readable ZIP entry.
The test accepts either parsing failure while still requiring the integrity
diagnostic; the particular parser exception is not the integrity assertion.
