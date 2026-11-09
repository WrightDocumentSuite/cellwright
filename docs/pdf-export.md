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

To stop instead, set `RefuseWhen`. It sees every `PdfExportWarning`, and the render raises
`PdfExportRefusedException` listing all the compromises it selects; no document is returned and a
named file is left as it was. `PdfExportRefusal` has ready-made policies: `Omissions`,
`OmissionsAndFontSubstitutions` and `AnyCompromise`.

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

System lookup reads the family, PostScript name, weight, and italic angle from the font file.
Consequently Office/PDF aliases and subset-prefixed names resolve independently of platform file
names such as Windows' `times.ttf`; regular, bold, italic, and bold-italic cells keep distinct faces
and their corresponding Excel print metrics. A Windows compatibility fallback for `times*.ttf`,
`cour*.ttf`, and `arial*.ttf` keeps this behaviour when Cellwright is consumed with an older pinned
Inkwright package.

## AutoFit

Only something that knows how wide text is can fit a column to it — and that is the font. AutoFit
therefore lives here, not in the core, and measures with exactly what printing uses: the same metrics
of the same face (including substituted), the same number format, the same padding around glyphs.
Because of that, a fitted column prints without clipping the text it was fitted for.

```csharp
Workbook book = await Workbook.LoadAsync("report.xlsx");

book.Sheets[0].AutoFit();                                  // columns, then rows
book.Sheets[0].AutoFitColumns(new AutoFitOptions { MaxWidth = 40 });
book.Sheets[0].AutoFitRows(new AutoFitOptions { IncludeMergedCells = true });
book.Sheets[0].Column("B").AutoFit();                      // one column
book.AutoFit();                                            // whole book, one font set for all sheets

await book.SaveAsync("report.xlsx");
```

The “columns, then rows” order is not cosmetic: the height of a wrapped cell depends on column width,
so rows must be measured after columns are settled. `AutoFit()` does that in the right order itself.

What does not participate in width fitting, and why:

- **Wrapped cell.** It sets its row height, not column width: fitting the column to it would undo the
  wrap the author asked for.
- **Merged cell.** Its text belongs to a rectangle, not one column. It never participates in width
  fitting. Row fitting also ignores it by default, matching Excel; set `IncludeMergedCells` to
  measure its text against the combined visible width of the merged range.
- **Hidden row.** A row that was hidden is not fitted to the grid. `IncludeHiddenRows` includes it
  again if that is still needed.

`AutoFitRows` differs from what printing measures: printing only raises a row with no declared height
— Excel already measured it and the author saw that number — while an explicit call sets height for
every row it measured. Calling AutoFit means asking to change them.

For a horizontal merge, merged-cell fitting sets the one row to the measured height. If a merge
spans several rows, their combined existing height is kept when it already fits; otherwise the
anchor row, then following rows, grow up to `MaxHeight` until the missing height is supplied. This
avoids flattening deliberate row proportions merely because a merge crosses them.

`SampleRows` limits measurement to the first N rows. On a twenty-thousand-row sheet that is the
difference between 112 ms and 31 ms, and the answer is usually decided by the first few hundred rows.

Limits: `MinWidth`/`MaxWidth` in characters (format ceiling 255), `MinHeight`/`MaxHeight` in points
(ceiling 409.5). `MaxWidth` defaults to the format ceiling, but for a report it is almost always
worth lowering: one long note otherwise pushes every other column off the sheet.

`Culture` must match the one the book will later print with — otherwise one string is measured and
another is drawn.

## Charts

Embedded charts are drawn: title, axes with round ticks and grid, category labels, legend, and the
series themselves — columns, bars, lines, areas, scatters (with X-Y numeric coordinate mapping and
point markers), radar spider-web polygons, pie, and doughnut. Points come from `numCache`/
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

A formula engine is needed for “expression” rules, comparisons whose operand is not a literal,
formula-valued graded thresholds, and an activity condition attached to a color scale, bar, or icon
set. The print package deliberately does not depend on `Cellwright.Formulas`, so the engine is
passed in from outside:

```csharp
var options = new PdfExportOptions
{
    RelativeFormulaEvaluator = (sheet, anchor, cell, formula) =>
        sheet.EvaluateCopied(formula, anchor, cell),
};

book.SaveAsPdf("report.pdf", options);
```

The copy-aware callback receives both the top-left cell of the rule's first range and the cell being
printed. `EvaluateCopied` compiles the stored formula text once and applies the relative row/column
offset while it runs, so `$D2`, `ROW()`, and `COLUMN()` retain Excel semantics without retaining one
translated program per target cell. The older `FormulaEvaluator` property remains available for
custom engines; it receives already-translated text and the target cell. The built-in
`EvaluateCopied` callback begins an independent formula pass per target, so volatile `NOW()` and
`TODAY()` are not frozen across the entire render; callers that require one shared volatile
timestamp must provide an engine callback with its own render-scoped session.

Without this, such a rule **is not applied and goes into diagnostics** with code
`conditional-format-formula`. Treating it as false silently is not allowed: an uncolored cell would
look like a deliberate author choice, not something print could not do.

How rules stack: by ascending `Priority`; the first matching rule has the highest priority and owns
every property it states. Lower-priority rules can still fill properties it left unstated — a rule
that sets only fill does not prevent a later rule from setting the font. `StopIfTrue` stops the walk.
Differential format replaces fill entirely, not blended: it names a solid color, so the cell’s own
pattern or gradient steps aside for the rule — as in Excel.
Border edges participate independently, so a higher-priority left edge and a lower-priority right
edge both appear. The current border model has no per-edge presence bit: `None` is treated as an
unspecified edge while rules are combined and therefore cannot explicitly clear one edge.

What is drawn, not just computed:

- **Color scale** replaces cell fill: it grades the cell itself, not a layer on top.
- **Data bar** — a bar from the left edge, interpolated between `MinimumLength` and `MaximumLength`
  as the value moves through the range. Their SpreadsheetML defaults are 10% and 90%.
- **Icon sets** are drawn vectorially, not as symbol-font glyphs: the font with the right glyphs is
  not on every machine, and a missing glyph is empty space where the sheet says something. The base
  format names seventeen sets; the x14 extension adds `3Stars`, `3Triangles`, and `5Boxes`. This
  renderer draws them with a family of shapes — circle, arrow, flag, triangle, star, boxes, diamond,
  check/cross/exclamation, scale, and circle fractions.
  An unknown set goes into diagnostics as `icon-set-<name>`, not drawn at random.
- `ShowValue = false` on an icon or data bar removes the number itself, leaving only the graphic.
- The icon takes space on the left; text starts after it, not under it.

A rule that changes font or number format participates in measurement too: a wrapped row grows for
the text that will be drawn, not the text without the rule.

A priority-bearing, non-custom x14 icon-set rule named `3Stars`, `3Triangles`, or `5Boxes` is
evaluated through an internal read-only projection while its source `extLst` remains authoritative
for saving. Custom `cfIcon` mixtures, x14 data-bar companions, and other rules the model does not
represent have nothing safe to apply and remain `conditional-format-opaque`.

**Cost.** Against the same styled 20,000×5 sheet, two comparison rules take 1.06× as long; a color
scale and data bar take 1.27×. Range statistics (min, max, percentiles, average, duplicate counts)
are taken once per rule and walk only occupied cells in the rule's range; overlapping range
references count a cell once. Format overlay on style is cached by the pair “style + format”: few
pairs, many cells. Formula-valued graded thresholds are evaluated once per rule/threshold, not once
per rendered cell.

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

**Rotation.** Quarter turns (90° and 180° in Excel's `TextRotation` encoding) and arbitrary diagonal
angles (1°–89° anticlockwise and 91°–179° clockwise) are drawn with an exact rotation matrix around the
cell's text box. Rotation 255 is letters stacked vertically, each under the previous.

**Merges.** A merged cell draws once over its whole area, even if its empty tail extends past the
used range.

**Decoration.** Row and column headings (`PrintHeadings`), cell fill, grid (`PrintGridLines`), and cell
borders — in that order so nothing paints over what came before. Pattern fill currently draws as the
pattern color. External hyperlinks become URI link annotations; internal workbook links (`Sheet!A1`,
`#Sheet!A1`, or defined names) become page-destination link annotations when the target sheet is
printed, or report `internal-link` when the target is not printed.

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

**Pictures and charts.** JPEG and PNG embed without recompression. A picture anchored to two cells stretches to
the actual grid; anchored to one keeps its size; anchored to nothing sits where it was placed on the
sheet. A picture crossed by a page break is drawn on both pages, each clipping to its edge — on the
first and last strip the edge is the paper itself, so a picture wider than its columns is not
clipped. Hidden pictures are not drawn; formats that cannot embed go into diagnostics. Embedded
charts (`Column`, `Bar`, `Line`, `Pie`, `Doughnut`, `Area`, `Scatter`, `Bubble`, `Radar`, and
combined Cartesian charts) are drawn into their anchor rectangles from cached or live worksheet
cells.

**Links.** The click area is an annotation, not a mark in the content stream, so page scale and
centering are baked into its rectangle: the area stays over its cell at any `Scale` and
`CenterHorizontally`.

## Interactive comments

An ordinary export still behaves like printing: classic notes and threaded review conversations
are omitted and reported as `comment` / `threaded-comments`. Opt in when the PDF should remain a
review document:

```csharp
var options = new PdfExportOptions { IncludeComments = true };
book.SaveAsPdf("review.pdf", options);
```

Every classic note and threaded root becomes a compact sticky-note annotation centred on the
top-right endpoint of the cell rectangle that was actually printed. A threaded reply becomes a PDF
reply through `/IRT`; the text, known workbook person, nullable UTC creation time and durable
threaded identifier are retained. A classic note has no source timestamp, so its PDF `/CreationDate`
and `/M` are absent rather than set to the export time. Several notes can share a cell, and a
repeated print-title cell emits its thread only once.

The icon is viewer content of a fixed size, not part of the grid's measured content. Enabling it
therefore changes neither cell layout nor page count. A comment outside the print rectangle, on a
hidden row/column, or on a swallowed non-anchor cell of a merge is not moved to a plausible-looking
place: it is omitted with `comment-anchors` or `comment-hidden`. Missing/ambiguous parents, cycles,
duplicate identifiers and cross-cell replies are handled without recursion and reported as
`comment-thread-dangling` / `comment-thread-malformed`; an unknown person leaves the PDF author
empty and produces `comment-authors`.

Excel records a resolved flag on each threaded message but not who resolved it or when. PDF review
`/State` requires the responsible user in `/T`, so the exporter does not fabricate an identity or
write an anonymous machine-readable `Completed` state. For every resolved source message it appends
an ordinary authorless, timeless informational reply to that message, with no `/State` or
`/StateModel`, and reports `comment-resolved-state`.

Mention ranges and opaque `extLst` markup on messages, collections, people and the person list have
no PDF annotation equivalent. Their aggregate counts are reported as `comment-metadata`; the opaque
XML itself is never copied into diagnostics.

## Tagged PDF

`PdfExportOptions.Tagged` builds a structure tree during draw:

| Excel | PDF |
| --- | --- |
| Sheet | `Table` |
| Row | `TR` |
| Cell | `TD` |
| Repeating header row cell | `TH` with Scope=Column |
| Picture | `Figure`, description goes in `/Alt` |
| Interactive note or reply | `Annot` with an `OBJR` |

Fills, grid, and borders are marked as artifacts, so every mark on the page either belongs to the
tree or states honestly that it belongs to no one. Document language comes from
`PdfExportOptions.Language`; a book with a title asks readers to show it instead of the file name.
Each interactive note and reply is inserted directly after its cell content in reading order. Its
`Annot` element owns the annotation through `OBJR`, while `StructParent` and `ParentTree` provide the
reverse link; pages carrying comments use structure tab order (`/Tabs /S`). Pop-up helper windows
are deliberately not structure content.

## What is not supported

These are explicit refusals, not gaps: each goes into diagnostics rather than turning into
something that looks right but is not. Diagnostics for the same issue speak once, so three sheets in
draft mode give one warning, not three.

- **Pivot tables.** In the book they round-trip through the file as-is, but print cannot draw them.
- **Chart straddling a page break** (`chart-split`) is not drawn: half a chart reads as a chart.
  Kinds the renderer does not know (`chart-Unknown`) are skipped with diagnostics.
- **Classic notes** (`comment`) and **threaded comments with replies** (`threaded-comments`) in the
  default print-like mode; set `IncludeComments` for interactive annotations.
- **Pattern fills** (`pattern-fill`) draw as the pattern’s solid color.
- **`Draft` and `BlackAndWhite`** do not affect output.
- **Header at scale** (`header-scale`) draws full size even when the sheet asks to scale it with the
  body; **`AlignHeaderWithMargins`** (`header-margins`) is not disabled.
- **Right-to-left sheet** (sheet-level `bidiVisual`): cell text direction is correct, but the grid
  itself is not mirrored.
- **Picture in header/footer** (`&G`).
- **Link to an unprinted sheet or unresolved name** (`internal-link`): when the target sheet is hidden
  or excluded by `SheetSelection`, the link has no destination page in the PDF.
- **Unknown paper code** (`paper-<code>`) prints on A4.

## Cost

BenchmarkDotNet 0.15.8, .NET 10.0.10, Windows Server 2022, Intel Core i7-8700; a 20,000-row sheet
over 5 columns — 100,000 cells and 445 pages. Measured 2026-08-01 with SDK 10.0.302:

| Scenario | Mean | Allocations |
| --- | ---: | ---: |
| No styles | 533.81 ms | 106.57 MB |
| Sparse print area, one occupied cell | 23.64 ms | 6.32 MB |
| With number formats | 642.05 ms | 124.72 MB |
| Same, with structure tree | 1,147.01 ms | 402.46 MB |
| Two comparison rules | 682.43 ms | 191.63 MB |
| Color scale and data bar | 816.88 ms | 203.65 MB |

The sparse case still creates all 445 pages named by `PrintArea`; it only avoids measuring and
painting default-blank cells. Before that fast path the same benchmark took 54.96 ms, versus
23.64 ms now (2.32× faster), with the same 6.32 MB of allocations. Number formats add roughly a
fifth to a dense sheet. Tagged output costs more because every cell becomes a structure element.
The main dense-sheet cost is measurement: every cell width is computed from real font metrics, and
a number is formatted before that.

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
