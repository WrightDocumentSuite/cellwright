# Export to Markdown

A workbook becomes a single `.md` file plus a folder of images: each sheet is a section with a
GitHub Flavored Markdown table, and cell values are rendered the way Excel displays them. The
conversion is pure — model to text — with no binary dependencies and no filesystem access, so it
lives in the `Cellwright` core next to the writers rather than in a separate package.

```csharp
using Cellwright;

Workbook book = await Workbook.LoadAsync("report.xlsx");

// Pure transformation: text and images in hand; place them wherever you need.
MarkdownDocument document = book.ToMarkdown();
Console.WriteLine(document.Text);
foreach (MarkdownImage image in document.Images)
    Console.WriteLine($"{image.FileName}: {image.Content.Length} bytes");

// Or straight to disk: workbook.md and media/ beside it.
await book.ExportMarkdownAsync("out/report");
```

## What becomes what

```csharp
var book = Workbook.Create();
book.Properties.Title = "Quarterly Report";
Worksheet sheet = book.Sheets.Add("Sales");
sheet["A1"] = "Product";
sheet["B1"] = "Amount";
sheet["C1"] = "Date";
sheet.Cell("B1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Right));

CellStyle money = CellStyle.Default.NumberFormat("#,##0.00");
CellStyle day = CellStyle.Default.NumberFormat("yyyy-mm-dd");
sheet["A2"] = "Coffee";
sheet.Cell("B2").SetValue(1234.5).SetStyle(money);
sheet.Cell("C2").SetValue(new DateTime(2026, 7, 25)).SetStyle(day);
sheet["A3"] = "Tea";
sheet.Cell("B3").SetValue(990).SetStyle(money);
sheet.Cell("C3").SetValue(new DateTime(2026, 7, 26)).SetStyle(day);
```

produces exactly this:

```markdown
# Quarterly Report

## Sales

| Product | Amount | Date |
| --- | ---: | --- |
| Coffee | 1,234.50 | 2026-07-25 |
| Tea | 990.00 | 2026-07-26 |
```

The level-one heading is `Properties.Title`, or `Workbook` when that is unset; override it with
`WorkbookTitle`. Each **visible** sheet then gets a `##` heading with its name and a table built
from `UsedRange`. A hidden sheet is omitted from the document — the same as in Excel. A sheet with
no content keeps its heading and a `(empty)` row, because a heading with no body reads like a
broken section.

Document lines are separated by `\n` regardless of platform, and the file ends with exactly one
newline: export must match byte for byte on any machine, or it cannot be kept in git.

## Values as Excel shows them, not as the file stores them

Every cell goes through the same `NumberFormatRenderer` that powers `TEXT` and cell display: a date
under `yyyy-mm-dd` becomes a date, `$#,##0.00` becomes a formatted amount, `0.0%` becomes a
percent. In the file those are all the same number; without the format, export would emit serial
days instead of dates.

**A formula exports its value, not its text.** A cell with `=SUM(B1:B9)` that displays `42` shows
`42` in the document. If the workbook arrived without calculated values, calculate it first —
`book.Calculate()` from the `Cellwright.Formulas` package.

**The default culture is invariant**, like the rest of the library's rendering: a document that
reads differently in Berlin than in Boston is a bug, not localization. For locale-specific number
formatting, use `new MarkdownExportOptions { Culture = CultureInfo.GetCultureInfo("de-DE") }`.

## Header row and alignment

GFM requires a header row: the separator line is mandatory, and **the number of cells in the header
must match the number of cells in the separator**, or the block is not a table at all (§4.10,
example 203). That is the one hard rule of the extension; both rows are written from the same
column count, not from whatever happens to appear in the data rows.

`HeaderRow` controls where the header comes from:

| Value | Header content | Data starts at |
| --- | --- | --- |
| `FirstRow` (default) | First row of `UsedRange` | Second row |
| `ColumnLetters` | Column letters `A`, `B`, `C`… | First row |

Column alignment comes from `Alignment.Horizontal` on the first-row cell — with `FirstRow` that is
the header cell itself; with `ColumnLetters` it is the cell that defines how the column looks.

| Cell alignment | Separator |
| --- | --- |
| `Left` | `:---` |
| `Center`, `CenterContinuous` | `:---:` |
| `Right` | `---:` |
| Everything else | `---` |

## Escaping

A backslash escapes any ASCII punctuation character (GFM §6.1), and escaping works inside table
cells (§4.10, example 200). Only characters that would break document structure are escaped:

- **Pipe** `|` becomes `\|` — otherwise the cell ends mid-value.
- **Backslash** `\` is doubled — otherwise it consumes the next character.
- **Line break** becomes `<br>` — a table cell cannot contain a block.
- **Leading and trailing spaces and tabs** become HTML entities — otherwise GFM trims them when
  parsing the table.

Everything else passes through as-is: a cell with `*note*` will render as italics. That is a
deliberate trade-off between fidelity and readability — a source file where every asterisk and
underscore is backslash-escaped is unreadable, and Markdown is chosen precisely for its source.
For literal text, escape values yourself before export.

`\r\n` becomes a single `<br>`, not two.

Workbook and sheet names are data, not ready-made Markdown markup: ASCII punctuation in headings
is escaped. So a workbook titled `*Q1*` reads as `*Q1*`, not as italic `Q1`.

## Hyperlinks

A cell with a link becomes `[text](target)`. Text is what the cell displays; if the cell is empty,
the `display` attribute stored beside the link is used, and if that is also missing, the target
itself. That order matches Excel, which shows the cell value and ignores the cached `display` on
read.

For an external link, the target is its `Target`; if a `Location` is stored beside it, it is
appended as `Target#Location`. For an internal link (`Sheet2!A1`), `Location` is written as-is:
there is no way to carry workbook navigation into the document, but dropping it would be worse.
Targets containing parentheses or spaces are wrapped in angle brackets — the form CommonMark keeps
for that case:

```markdown
| [Example](https://example.com/a) | [Article](<https://example.com/Article_(disambiguation)>) |
```

Potentially active schemes such as `javascript:`, `vbscript:`, and `data:`, empty targets, and
targets with control characters are not made clickable: export leaves only the cell's visible text.

## Merged cells

GFM has no `colspan`. The merged range's value is written in its top-left cell; the rest of the
range stays empty — the same as on the sheet:

```markdown
| Total |  |
```

## Pictures

Sheet pictures are collected in an `### Images` section under the sheet's table, linked to files
in `media/`. They cannot go inside cells: a GFM table holds neither merged spans nor floating
objects.

```markdown
### Images

![Process diagram](media/image1.png)
```

- **Order follows anchors top to bottom**, not z-order: the document is read from the start, and
  figure stacking on the page does not matter.
- **Alt text** is the picture's `Description`, then `Name`, then the file name. Description comes
  first because it is the file's alt text; the name is usually how Excel numbered the shape.
- **A hidden picture is omitted**, like a hidden sheet.
- **The file extension follows the picture format** (`DrawingWriter.Extension`), not a name the
  shape does not have.
- **Content deduplication (SHA-256)**: the same logo on three sheets is one file in `media/` and
  three links. Bytes are written as-is, without re-encoding, as everywhere else in the library.

`ToMarkdown` writes nothing to disk: `MarkdownDocument.Images` returns file names and bytes, and
you can place them on a CDN, in object storage, or in a response stream. `SaveAsync` (and
`ExportMarkdownAsync` on top of it) is a convenience layer that writes `workbook.md` and `media/`
to a directory.

`MediaDirectoryName` must always be relative: absolute paths, `..`, and values without a directory
name are rejected. Backslashes are normalized to `/`, and each path segment is encoded separately
for links, so a directory `assets/quarter #1` is saved under that name and gets the correct URL
`assets/quarter%20%231/image1.png`.

## Options

| Option | Default | Effect |
| --- | --- | --- |
| `Culture` | `InvariantCulture` | Separators and month names when rendering cells |
| `HeaderRow` | `FirstRow` | Where the required table header comes from |
| `IncludePictures` | `true` | Whether to collect the image gallery and files |
| `MediaDirectoryName` | `media` | Safe relative directory that links point to |
| `WorkbookTitle` | — | Level-one heading instead of `Properties.Title` |

## Deliberate limitations

- **Export only.** There is no Markdown import back into a workbook.
- **Formatting does not carry over.** Colors, fills, fonts, borders, and conditional formatting
  cannot be expressed in Markdown; of styling, only number format and horizontal alignment survive.
- **Column widths, row heights, frozen panes, and grouping** are not exported: GFM tables have
  none of these.
- **Merges lose span**, not value — see above.
- **Comments, data validation, and auto-filters** are not included in the document.
- **The workbook is built entirely in memory.** Pictures still need a sidecar, so streaming text
  alone would gain nothing; a sheet with a million rows produces a string of corresponding size.
- **Internal links do not work as navigation** — they are carried over as text.

## How this is verified

The project has no Markdown parser and therefore no oracle, so verification rests on two legs:
exact strings and structural rules of the format.

- **Golden tests**: a workbook is built in code and the document is compared to the full expected
  text — headings, table, separator, alignment, gallery. The main example scenario is pinned by a
  separate test; when the example changes, both checks are updated.
- **GFM rule**: for every table, header cell count is computed the same way a format reader would
  (an escaped `|` is content, not a delimiter) and compared to the separator cell count.
- **Escaping**: cells with `|`, `\`, `\r\n`, edge whitespace, and Markdown-significant symbols;
  literal headings and unsafe links are checked separately.
- **Formats**: date, currency, and percent against what Excel shows; a formula exports its value and
  its text never appears in the document; the same workbook under `de-DE` changes separators.
- **Pictures**: gallery in place, files written, duplicates deduplicated, bytes match source,
  hidden ones skipped, order by anchor.
- **Directory**: `ExportMarkdownAsync` writes `workbook.md` and `media/`; a workbook without
  pictures gets no `media/` directory at all.

## Reference documents

| Document | What we follow from it |
| --- | --- |
| GitHub Flavored Markdown Spec §4.10 (Tables) | Table structure, separator, alignment, pipe escaping |
| GFM §6.1 (Backslash escapes) | What is escaped in cell text and how |
| GFM §6.6, §6.7 (Links, Images) | Links, images, angle brackets around targets |
