# PDF export

`Cellwright.Pdf` prints a workbook the way Excel prints it: a grid with its own styles and number
formats, cut into pages that fit the paper. Rendering is done by
[Inkwright](https://github.com/WrightDocumentSuite/inkwright) — no Office, no native dependencies —
so it behaves the same on a build agent as on a desktop.

Inkwright is an ordinary NuGet dependency, pinned to the exact version this package was built and
tested against; nothing outside the repository has to be checked out to build from source.

```csharp
using Cellwright;
using Cellwright.Pdf;

Workbook book = await Workbook.LoadAsync("report.xlsx");
PdfExportDiagnostics diagnostics = book.SaveAsPdf("report.pdf");

foreach (PdfExportWarning warning in diagnostics)
    Console.WriteLine(warning);
```

Rendering never throws because of content it cannot draw. It draws what it can and records the rest
in diagnostics — a batch job continues, and a person can later see which files deserve a visual
check.

## Get a document instead of saving

`SaveAsPdf` is the short path. When the result must be signed, encrypted, or declared archival, the
document is returned open:

```csharp
using Inkwright;

PdfExportResult result = book.ToPdf(new PdfExportOptions { Tagged = true });
using PdfDocument pdf = result.Document;

pdf.Save("report.pdf");
```

One sheet prints on its own: `sheet.ToPdf()`.

## Which sheets print

By default — all visible sheets, in tab order. `PdfExportOptions.Selection` changes this to every
sheet in order (`All`) or a named list (`Named` plus `Sheets`). A sheet that is not in the book goes
into diagnostics, not a failed export.

## Units

Column width in the file is not a length but a count of digits in the book’s standard font, so its
size in points depends on how wide that font draws a digit. The conversion is documented by the
format (ISO/IEC 29500-1 §18.3.1.13): pixels at 96 dpi first, then points at 72. The same file under
a different standard font prints slightly differently — and Excel behaves the same way. Page margins
are stored in inches; paper size by code.

All paper codes listed in §18.3.1.63 are supported: from Letter and Ledger (17×11 — Tabloid on its
side, not the same thing) through A2, Japanese envelopes, and rotated sizes. A code not in the list
prints on A4 and is reported in diagnostics; code 0 means “printer default” and also yields A4, but
silently, because no choice was made here.

## Fonts

The book names a family; the machine finds the file. That works on a desktop, not on a build agent,
so lookup follows a chain, and each step past the requested family is recorded in diagnostics as
`FontSubstituted`:

1. files registered by the caller (`PdfExportOptions.FontFiles`),
2. fonts installed on the system,
3. a table of metrically or visually close faces — Carlito for Calibri, Liberation Serif for Times
   New Roman, and so on,
4. `PdfExportOptions.FallbackFontFamily`,
5. one of fourteen fonts any reader carries — chosen from what the name suggests.

```csharp
var options = new PdfExportOptions();
options.FontFiles["Calibri"] = "/srv/fonts/Carlito-Regular.ttf";
options.FontFiles["Calibri Bold"] = "/srv/fonts/Carlito-Bold.ttf";
```

## AutoFit

Only something that knows how wide text is can fit a column to it — and that is the font. AutoFit
therefore lives here, not in the core, and measures with exactly what printing uses: the same metrics
of the same face (including substituted), the same number format, the same padding around glyphs.
Because of that, a fitted column prints without clipping the text it was fitted for.

```csharp
Workbook book = await Workbook.LoadAsync("report.xlsx");

book.Sheets[0].AutoFit();                                  // columns, then rows
book.Sheets[0].AutoFitColumns(new AutoFitOptions { MaxWidth = 40 });
book.Sheets[0].Column("B").AutoFit();                      // one column
book.AutoFit();                                            // whole book, one font set for all sheets

await book.SaveAsync("report.xlsx");
```

The “columns, then rows” order is not cosmetic: the height of a wrapped cell depends on column width,
so rows must be measured after columns are settled. `AutoFit()` does that in the right order itself.

What does not participate in width fitting, and why:

- **Wrapped cell.** It sets its row height, not column width: fitting the column to it would undo the
  wrap the author asked for.
- **Merged cell.** Its text belongs to a rectangle, not one column. Excel also leaves merges out of
  AutoFit.
- **Hidden row.** A row that was hidden is not fitted to the grid. `IncludeHiddenRows` includes it
  again if that is still needed.

`AutoFitRows` differs from what printing measures: printing only raises a row with no declared height
— Excel already measured it and the author saw that number — while an explicit call sets height for
every row it measured. Calling AutoFit means asking to change them.

`SampleRows` limits measurement to the first N rows. On a twenty-thousand-row sheet that is the
difference between 112 ms and 31 ms, and the answer is usually decided by the first few hundred rows.

Limits: `MinWidth`/`MaxWidth` in characters (format ceiling 255), `MinHeight`/`MaxHeight` in points
(ceiling 409.5). `MaxWidth` defaults to the format ceiling, but for a report it is almost always
worth lowering: one long note otherwise pushes every other column off the sheet.

`Culture` must match the one the book will later print with — otherwise one string is measured and
another is drawn.

## Charts

Embedded charts are drawn: title, axes with round ticks and grid, category labels, legend, and the
series themselves — columns, bars, lines, areas, pie, and doughnut. Points come from `numCache`/
`strCache` caches carried by the file; a series without caches is read from live cells through a
reference. Series colors are book theme accents, in rotation.

A chart is drawn on the page where it fits whole; one straddling a break is not drawn at all and
goes into diagnostics (`chart-split`) — half a chart reads as a chart. Axis titles, line styles,
rotated labels, and the rest of the richness of §21.2 are not reproduced: data is drawn, not
decoration.

## Conditional formatting

Rules are evaluated, not ignored: a cell prints as Excel shows it.

Almost all rule kinds work without a formula engine — comparisons to literals, text tests (`contains`,
`begins with`, `ends with`), blank and error cells, duplicates and unique values, Top 10 and
top/bottom percent, above/below average with deviations, all ten date periods from `ST_TimePeriod`, color
scales, data bars, and icon sets. That is exactly what is decided from values, and print values are
enough.

A formula is needed only for “expression” rules and comparisons whose operand is not a literal.
The print package deliberately does not depend on `Cellwright.Formulas`, so the engine is passed in
from outside:

```csharp
var options = new PdfExportOptions
{
    FormulaEvaluator = (sheet, cell, formula) => sheet.Evaluate(formula, cell.ToString()),
};

book.SaveAsPdf("report.pdf", options);
```

Without this, such a rule **is not applied and goes into diagnostics** with code
`conditional-format-formula`. Treating it as false silently is not allowed: an uncolored cell would
look like a deliberate author choice, not something print could not do.

How rules stack: by ascending `Priority`, each next rule overrides the previous only where it says
something — a rule that sets only fill leaves the font chosen earlier. `StopIfTrue` stops the walk.
Differential format replaces fill entirely, not blended: it names a solid color, so the cell’s own
pattern or gradient steps aside for the rule — as in Excel.

What is drawn, not just computed:

- **Color scale** replaces cell fill: it grades the cell itself, not a layer on top.
- **Data bar** — a bar from the left edge, as long a fraction of the cell as the value is of the
  range.
- **Icon sets** are drawn vectorially, not as symbol-font glyphs: the font with the right glyphs is
  not on every machine, and a missing glyph is empty space where the sheet says something. The format
  names seventeen sets (`ST_IconSetType`); this renderer draws them with a family of shapes — circle,
  arrow, flag, triangle, diamond, check/cross/exclamation, scale, and circle fractions.
  An unknown set goes into diagnostics as `icon-set-<name>`, not drawn at random.
- `ShowValue = false` on an icon or data bar removes the number itself, leaving only the graphic.
- The icon takes space on the left; text starts after it, not under it.

A rule that changes font or number format participates in measurement too: a wrapped row grows for
the text that will be drawn, not the text without the rule.

A rule the model does not represent (for example with x14 extensions) is carried through the file
verbatim, but print has nothing to apply — `conditional-format-opaque`.

**Cost.** On a 20,000×5 sheet with two comparison rules — 1.66× the time of a sheet without rules;
with a color scale and data bar — 1.77×. Range statistics (min, max, percentiles, average, duplicate
counts) are taken once per rule, not per cell; otherwise a column of a thousand values would be
quadratic. Format overlay on style is cached by the pair “style + format”: few pairs, many cells.

## What prints

**Values.** Each passes through the same number-format renderer as `TEXT`: sections, dates,
fractions, percents, conditions. Culture comes from `PdfExportOptions.Culture` — invariant by default
so the document looks the same on any machine. Formulas print their cached value; calculating the
book before print is the caller’s job (`book.Calculate()`).

**Alignment.** Horizontal — left, right, center, and `General`, which looks at type: numbers right,
text left, logical and error centered. Vertical — top, middle, bottom. Indent shifts text from the
edge, as in Excel.

**Width and overflow.** A number wider than its column prints as hash marks — a hint to widen the
column, not an error. Text behaves differently: without wrap it spills into neighboring **empty**
cells (right for left align, left for right, both ways for center) and clips at the first neighbor
that has content. What the reader will not see is not drawn at all: clipping does not hide glyphs
behind a clip path but omits them, otherwise search would find text not on the page.

**Wrap and row height.** `wrapText` wraps on word boundaries; a word wider than the cell is cut where
space runs out. A row with no height in the file grows for wrapped lines; a row with a set height
keeps it, and extra lines simply do not print — Excel does the same.

**Rotation.** Quarter turns in either direction use one matrix; intermediate angles — nearest quarter
with diagnostics. Rotation 255 is letters stacked vertically, each under the previous.

**Merges.** A merged cell draws once over its whole area, even if its empty tail extends past the
used range.

**Decoration.** Cell fill, grid (`PrintGridLines`), and cell borders — in that order so nothing
paints over what came before. Pattern fill currently draws as the pattern color. An external
hyperlink becomes an annotation.

**Pages.** The grid is cut in both directions at once: columns into strips by paper width, rows by
paper height, and each strip pair is a page. Manual breaks are hard edges. Repeating rows and columns
(`SetRepeatRows`, `SetRepeatColumns`) print on every page and reduce space for the rest. Walk order —
`PageOrder`; centering — `CenterHorizontally` and `CenterVertically`.

**Scale.** `Scale` applies as given; `FitToWidth`/`FitToHeight` pick the largest scale where strip
count does not exceed the request. The page body draws under one scale matrix — headers and footers
outside it, in their own field.

**Headers and footers.** Format codes parse (§18.3.1.46): `&L&C&R` sections, page number `&P`
(including offset `&P+2`), page count `&N`, sheet name `&A`, date and time, font codes
`&"Arial,Bold"`, `&14`, `&B&I&U&S`, color `&K`. A font code applies to what follows and does not
reach back: in `small&20BIG` only `BIG` prints at twenty points. Separate headers for the first page
and for even pages — by flags `DifferentFirstPage` and `DifferentOddEven`; empty text with the flag
set means a page without that header, not “use the usual one”. Each sheet counts its own pages: `&P`
shows the number within the sheet, as in Excel.

**Pictures.** JPEG and PNG embed without recompression. A picture anchored to two cells stretches to
the actual grid; anchored to one keeps its size; anchored to nothing sits where it was placed on the
sheet. A picture crossed by a page break is drawn on both pages, each clipping to its edge — on the
first and last strip the edge is the paper itself, so a picture wider than its columns is not
clipped. Hidden pictures are not drawn; formats that cannot embed go into diagnostics.

**Links.** The click area is an annotation, not a mark in the content stream, so page scale and
centering are baked into its rectangle: the area stays over its cell at any `Scale` and
`CenterHorizontally`.

## Tagged PDF

`PdfExportOptions.Tagged` builds a structure tree during draw:

| Excel | PDF |
| --- | --- |
| Sheet | `Table` |
| Row | `TR` |
| Cell | `TD` |
| Repeating header row cell | `TH` with Scope=Column |
| Picture | `Figure`, description goes in `/Alt` |

Fills, grid, and borders are marked as artifacts, so every mark on the page either belongs to the
tree or states honestly that it belongs to no one. Document language comes from
`PdfExportOptions.Language`; a book with a title asks readers to show it instead of the file name.

## What is not supported

These are explicit refusals, not gaps: each goes into diagnostics rather than turning into
something that looks right but is not. Diagnostics for the same issue speak once, so three sheets in
draft mode give one warning, not three.

- **Pivot tables.** In the book they round-trip through the file as-is, but print cannot draw them.
- **Chart straddling a page break** (`chart-split`) is not drawn: half a chart reads as a chart.
  Kinds the renderer does not know (`chart-Unknown`) are skipped with diagnostics.
- **Notes and comments** (`comment`) on print.
- **Row and column headings** (`print-headings`) when the sheet asks to print them.
- **Pattern fills** (`pattern-fill`) draw as the pattern’s solid color.
- **Intermediate rotation angles** (`rotation`) round to the nearest quarter turn.
- **`Draft` and `BlackAndWhite`** do not affect output.
- **Header at scale** (`header-scale`) draws full size even when the sheet asks to scale it with the
  body; **`AlignHeaderWithMargins`** (`header-margins`) is not disabled.
- **Right-to-left sheet** (sheet-level `bidiVisual`): cell text direction is correct, but the grid
  itself is not mirrored.
- **Picture in header/footer** (`&G`).
- **Link within the book**: in a stack of pages it has nowhere to point.
- **Unknown paper code** (`paper-<code>`) prints on A4.

## Cost

BenchmarkDotNet 0.15.8, .NET 10, Windows Server 2022, Intel Core i7-8700; a 20,000-row sheet over
5 columns — 100,000 cells, 445 pages:

| | Mean | Allocations |
| --- | ---: | ---: |
| No styles | 827 ms | 106 MB |
| With number formats | 1,014 ms | 125 MB |
| Same, with structure tree | 1,728 ms | 333 MB |

Under ten microseconds per cell; number formats add roughly a fifth; the structure tree another
two-thirds and triple the memory, because every cell becomes an element. The main cost is
measurement, not output: every cell width is computed from real font metrics, and a number is
formatted before that.

```bash
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks -- --filter *PdfBenchmarks*
```

## How it works

Print runs in two passes per sheet: the grid is measured, then walked — separated so the second pass
can repeat at another scale without measuring again.

```
SheetPlan      what prints: print area, hidden rows/cols out, repeating headers
GridMeasurer   digits → points, auto row heights with wrap
CellPlacer     value → text, alignment, spill, clip, rotation
SheetComposer  grid → page strips, scale, headers/footers
PageRenderer   page elements → content stream operators
```

Everything before the renderer works in points with the origin at the top-left — as Excel thinks.
The flip into PDF coordinates happens in one place, `PageGeometry.ToPdfY`.

The page body draws under one matrix — sheet scale and where the grid landed on the paper. The
matrix applies to the content stream but not annotations, so link areas are computed by the renderer
itself (`ItemPainter.EnterBody`). The `PdfExportOptions.MaxPages` limit is checked by strip count
before pages are built: a limit noticed after the fact would not save the memory it exists for.
