# Worksheet: rows, columns, and everything around cells

Besides values, a worksheet carries what people usually open Excel for: hidden rows, grouping,
auto-filter, links, comments, tables, validation rules, conditional formatting, print settings, and
protection. All of that is modeled; anything that is not is preserved through load and save — see
[Preserving foreign files](#preserving-foreign-files).

## Rows and columns

Facades like `Cell`: a sheet + index pair; mutators return `this` for chaining.

```csharp
sheet.Row(1).SetHeight(28);
sheet.Row(3).Hide();
sheet.Column("B").SetWidth(18).Hide();

sheet.Rows(4, 12).Group();                 // outline level +1
sheet.Columns(3, 5).Group(collapse: true); // group and collapse immediately
sheet.Rows(4, 12).Show();

bool hidden = sheet.Row(3).Hidden;
double? width = sheet.Column(2).Width;     // null — default width
```

`ConfiguredRows` and `ConfiguredColumns` enumerate only what differs from defaults: an empty sheet
does not store a million entries. Adjacent identical columns merge into one `<col min max>` on
write, so `sheet.Columns(1, 100).SetWidth(9)` is one element in the file, not a hundred. Width is
capped at 255 characters: Excel refuses to open a file with more.

AutoFit from here is impossible: only a font knows how wide text is. AutoFit therefore lives in
`Cellwright.Pdf` — `sheet.AutoFit()`, see [pdf-export.md](pdf-export.md#autofit).

Hidden rows are honored by `SUBTOTAL` with codes 101–111 and `AGGREGATE` with options 1/3/5/7 —
see [formulas.md](formulas.md).

## Inserting and deleting rows and columns

```csharp
sheet.InsertRows(2, 3);      // three empty rows above row 2
sheet.DeleteRows(10);        // delete row 10
sheet.InsertColumns(4);
sheet.DeleteColumns(2, 2);
```

More than cells move. Row heights and column widths, grouping, merges, hyperlinks, comments,
validation and conditional-formatting rules, table ranges, auto-filter, page breaks, and sheet-
protection exceptions follow. Formulas across the whole workbook are retargeted: a range that gains
a row inside it expands; a reference to a deleted cell becomes `#REF!`; defined names (including
print area) are updated the same way. Absolute references shift together with relative ones — `$`
protects against copying a formula, not against the grid moving underneath.

```csharp
// Was: C1 = SUM(B1:B5), D1 = B3*2
sheet.InsertRows(2);       // C1 = SUM(B1:B6), D1 = B4*2
sheet.DeleteRows(4);       // C1 = SUM(B1:B5), D1 = #REF!*2
```

A new row takes the height and outline level of the row above it, but not its hidden state: a row
born invisible would be surprising. Cells pushed past the grid edge are dropped — Excel refuses the
insert in that case; here the overflow simply disappears.

What is lost: saved worksheet markup with coordinates the model does not represent — sort state
(`sortState`), custom views, and “ignore error” flags (`ignoredErrors`). They are dropped rather
than left pointing at the wrong place; a warning arrives in `Workbook.OnWarning` with code
`StructuralEditDroppedMarkup`. Picture and chart anchors live in their own part and stay at their
old coordinates.

## Worksheet appearance

```csharp
sheet.TabColor = ExcelColor.FromRgb(0x22, 0x78, 0xD4);
sheet.ShowGridLines = false;
sheet.ZoomScale = 145;
sheet.FrozenRows = 1;
sheet.Visibility = SheetVisibility.VeryHidden;
```

Visibility is three-way, as in the spec: `Hidden` is restored through the sheet tab context menu;
`VeryHidden` only from the VBA editor — what files use to hide service sheets seriously. The
`IsHidden` property remains for the simple case and sets ordinary `Hidden`.

## Hyperlinks

Excel stores them two different ways: external links as worksheet relationships, internal ones as
a `location` attribute. The model uses one type; `IsExternal` tells them apart.

```csharp
sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com", tooltip: "Open site");
sheet.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("A3"), "Summary!A1"));

Hyperlink? link = sheet.Hyperlinks.Find(CellRef.Parse("A2"));
sheet.Hyperlinks.Remove(RangeRef.Parse("A2:A3"));
```

## Auto-filter

```csharp
sheet.AutoFilter("A1:D100");
RangeRef? range = sheet.AutoFilterRange;
sheet.RemoveAutoFilter();
```

Filter criteria themselves are not modeled. In a loaded file they are preserved and written back
while the range is unchanged; after moving the filter only the drop-down lists remain — a column
condition that no longer sits in the same place would hide the wrong rows.

## Comments

Classic yellow notes (not threaded comments). Text is flat: rich text is flattened on read.

```csharp
sheet.Comments.Add("B2", "Check the number", author: "Auditor");
CellComment? note = sheet.Comments[CellRef.Parse("B2")];
sheet.Comments.Remove(CellRef.Parse("B2"));
```

While comments are untouched, the `comments` and VML parts are copied byte for byte — yellow box
positions and sizes stay as they were. The first edit regenerates both parts with default geometry:
a box two columns to the right of the cell and four rows tall, like a fresh note in Excel. Form
controls that lived in the same VML are carried verbatim: a checkbox is not the price of fixing a
typo in a comment. New boxes are numbered after preserved shapes, and `o:idmap` declares every
occupied block — a duplicate id would send the file into recovery. Even after removing the last
comment, the sheet keeps its VML while foreign shapes remain.

Geometry is anchored with `x:Anchor`, not absolute shape position — that is how Excel places the
box next to its cell. Shape ids are issued in blocks of 1024 per drawing — one block per sheet —
otherwise two sheets with comments would declare the same `o:idmap` and Excel would treat the file
as needing repair.

## Threaded comments

Discussions with author and replies that Excel introduced instead of yellow notes. This is an
extension, not part of ISO/IEC 29500: parts, content types, and relationships are described in
MS-XLSX 2.1.17–2.1.18; elements — in 2.6.203–2.6.207.

```csharp
Person nadia = book.People.Add("Nadia");
Person tom = book.People.Add("Tom");

ThreadedComment thread = sheet.ThreadedComments.Add(CellRef.Parse("B2"), "Net or gross?", nadia);
sheet.ThreadedComments.Reply(thread, "Net.", tom);
thread.Resolved = true;

foreach (ThreadedComment root in sheet.ThreadedComments.Threads)
    Console.WriteLine($"{root.Cell}: {root.Text} (+{root.Replies.Count})");
```

There are two parts at two levels: threads hang off the sheet; people hang off the workbook,
because an author who writes on three sheets is one person. A comment names its author by id;
`book.People` resolves it.

Nesting is one level deep: a reply to a reply attaches to the thread, not a chain — that is how
the format works and how Excel shows the discussion. `Threads` builds a tree from a flat message
list; `Remove` on the first message takes its replies with it.

**Stub note.** Next to each thread the file places a classic note with an author like `tc={id}` so
a reader that knows nothing about threads still shows something (MS-XLSX 2.3.7.3). Such stubs do
not appear in `Comments` — their author is an id, not a person — and on save they are
**regenerated from the threads**. Editing ordinary comments therefore no longer detaches threads
from cells: previously rewriting the `comments` part removed the stubs and a `ThreadedComments`
warning said so honestly; that warning is gone because there is nothing left to lose.

The converse follows: editing threads, comments, or the people list regenerates both parts
together. While none of that is touched, the file carries them byte for byte.

Not modeled: `mentions` (people mentioned inside text) and `extLst`; they are lost when a thread is
rewritten.

## Excel tables

```csharp
ExcelTable table = sheet.Tables.Add("Sales", "A1:D100");
table.Range = RangeRef.Parse("A1:D150");     // expand

sheet.Tables.Add(new ExcelTable
{
    Name = "Lookup",
    Range = RangeRef.Parse("F1:G20"),
    Columns = ["Code", "Name"],
    StyleName = "TableStyleLight9",
    ShowRowStripes = true,
});
```

Column names come from the header row; empty header cells become `Column2`, `Column3`, and so on,
as in Excel. Table filter state and calculated-column formulas are not modeled: an untouched table
keeps them in its part; any change to the sheet’s table list regenerates parts from the model.
Structural references `Table[Column]` are not yet understood by the formula engine.

## Data validation

```csharp
sheet.Validations.Add(DataValidation.List("E2:E100", "\"Yes,No\""));
sheet.Validations.Add(DataValidation.List("F2:F100", "Lookups!$A$1:$A$20"));
sheet.Validations.Add(DataValidation.WholeBetween("G2:G100", 1, 99));
sheet.Validations.Add(DataValidation.Custom("H2:H100", "LEN(H2)<=10"));

sheet.Validations.Add(new DataValidation
{
    Ranges = [RangeRef.Parse("B2:B100"), RangeRef.Parse("D2:D100")],
    Type = DataValidationType.Decimal,
    Operator = DataValidationOperator.GreaterThan,
    Formula1 = "0",
    ErrorTitle = "Positive only",
    Error = "Enter a number greater than zero.",
    ErrorStyle = DataValidationErrorStyle.Warning,
});
```

Extended Excel rules (`x14`, lists from another sheet) live in the sheet’s `extLst` and survive
load-save while rules are unchanged.

## Conditional formatting

Rules are read, edited, and created anew. Order is by `Priority`: lower numbers are tried first;
`Add` assigns the next free priority itself.

```csharp
sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B2:B99", 1000, DifferentialStyle.Good));
sheet.ConditionalFormats.Add(ConditionalFormatRule.CellBetween("B2:B99", 0, 100, DifferentialStyle.Bad));
sheet.ConditionalFormats.Add(ConditionalFormatRule.Containing("C2:C99", "urgent", DifferentialStyle.Neutral));
sheet.ConditionalFormats.Add(ConditionalFormatRule.Duplicates("D2:D99", DifferentialStyle.Bad));
sheet.ConditionalFormats.Add(ConditionalFormatRule.Top("B2:B99", 3, DifferentialStyle.Good));

// Write the formula as it would be read in the top-left cell of the range:
// highlight the whole row from one column’s value
sheet.ConditionalFormats.Add(
    ConditionalFormatRule.Expression("A2:F99", "$D2>100", DifferentialStyle.Good));

// Gradients, bars, and icons
sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("B2:B99", ColorScale.RedYellowGreen));
sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("E2:E99", new DataBar { ShowValue = false }));
sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("F2:F99", IconSet.TrafficLights));
```

A rule colors with a **differential format** (`dxf`) — overrides on top of the cell’s own formatting
— not a full cell style. Only what is listed is set:

```csharp
var overdue = new DifferentialStyle
{
    Bold = true,
    FontColor = ExcelColor.FromRgb(0x9C, 0x00, 0x06),
    FillColor = ExcelColor.FromRgb(0xFF, 0xC7, 0xCE),
    NumberFormat = "0.00%",
};
```

Ready-made `DifferentialStyle.Good`, `.Bad`, and `.Neutral` mirror Excel’s presets. Formats are
interned: two rules with the same `DifferentialStyle` share one `dxf`.

A rule read from a file exposes its format through the same property: the file’s `dxf` table is
parsed and `rule.Style` returns bold, colors, number format, and borders the rule applies. Until
the style is rewritten, the rule keeps pointing at the original format in the saved table — it is
richer than the parsed view (alignment, for example), and there is no reason to drop it after a
read. Assigning `rule.Style` is an ordinary edit: the sheet switches to regenerating rules and the
new format is appended after preserved ones.

While sheet rules are untouched, `conditionalFormatting` blocks return to the file as they were.
The first edit regenerates them from the model; rules the model does not describe (modern icon sets
and everything in the `x14` extension) return as their original XML — they have `IsOpaque` true and
`dxfId` still points into the saved format table.

## Printing

```csharp
sheet.PageSetup.Orientation = PageOrientation.Landscape;
sheet.PageSetup.PaperSize = 9;                   // A4
sheet.PageSetup.FitToWidth = 1;                  // fit width; height free
sheet.PageSetup.Margins = PageMargins.Narrow;
sheet.PageSetup.PrintGridLines = true;
sheet.PageSetup.CenterHorizontally = true;
sheet.PageSetup.Header = new HeaderFooterText { Default = "&CQuarterly report" };
sheet.PageSetup.Footer = new HeaderFooterText { Default = "&RPage &P of &N" };
sheet.PageSetup.AddRowBreak(30);                 // page break above row 30

sheet.SetPrintArea("A1:H40");
sheet.SetRepeatRows(1, 2);                       // header on every page
sheet.SetRepeatColumns(1, 1);
```

Header and footer codes work like Excel: `&L`/`&C`/`&R` split a line into sections, `&P` is page
number, `&N` is page count, `&A` is sheet name, `&D` is date. Print area and repeating rows are
stored not on the sheet but in built-in names `_xlnm.Print_Area` and `_xlnm.Print_Titles`; these
properties read and write them.

While `PageSetup` is untouched, original elements return as they were — including a link to a
printer-settings part the model does not represent.

## Sheet order and copying

```csharp
book.Sheets.Move("Summary", 0);                    // to first position
Worksheet copy = book.Sheets.Copy("January", "February");   // copy inserted right after original
```

A copy gets everything the model represents: values, styles, formulas, widths and heights, merges,
freeze panes, auto-filter, comments, links, validation rules, conditional formatting, print
settings, and protection. Every mutable object is copied by value, not shared: inserting a row in
the copy does not shift a rule on the original.

A copy does not get what belongs to the source worksheet part in the package: drawings, pictures,
and other raw markup. One part cannot live at two paths, and a copy is a new sheet, not a second
name for the old one. Formulas are copied as text, so in-sheet references follow the copy while
those that name the source sheet by name still point at it. Table names are unique in the
workbook, so a copied table gets its own: `Sales` → `Sales2`.

## Document properties

What the file says about itself — Excel’s Info tab. A workbook created from scratch now ships with
the same `docProps` parts Excel writes, not without them.

```csharp
book.Properties.Title = "Quarterly report";
book.Properties.Author = "Anna Smith";
book.Properties.Company = "Contoso";
book.Properties.Custom["Department"] = CellValue.Text("Finance");
book.Properties.Custom["Approved"] = CellValue.Boolean(true);
```

There are `Title`, `Subject`, `Author`, `Keywords`, `Description`, `Category`, `ContentStatus`,
`LastModifiedBy`, `Created`, `Modified`, `Application`, `Company`, `Manager`, and arbitrary
`Custom` fields — text, number, flag, or date, the four types the format allows.

`Modified` updates on every save unless you set it yourself; a value you set is written as-is.
Fields the model does not name — language, revision, identifier, Word statistics in `app.xml` —
return to the file as they arrived: both parts allow children in any order, so appending generated
fields is safe. Exception — two sheet-name lists (`HeadingPairs`, `TitlesOfParts`): the model
rebuilds them because sheets may have been renamed.

## Protection

Protection in xlsx is not encryption: anyone can read the data; a password only stops editing in
Excel. The model reflects that — it stores a verifier, not the password.

```csharp
sheet.Protection.Protect("hunter2");
sheet.Protection.AllowSort = true;               // what remains allowed
sheet.Protection.AllowAutoFilter = true;
sheet.Protection.EditableRanges.Add(ProtectedRange.For("Entry", "B2:D100"));

book.Protection.Protect();                       // forbid changing the sheet list

bool ok = sheet.Protection.Password?.Matches("hunter2") ?? true;
```

The modern verifier is written: SHA-512 with salt and 100,000 iterations. The legacy 16-bit
checksum (pre–Excel 2013 files) is read and verified but never recreated — it can be brute-forced
in seconds.

## Preserving foreign files

Load remembers everything the model does not represent; save writes it back byte for byte: charts,
pivot tables and their caches, pictures, themes, `.xlsm` macros, form controls, chart sheets.
Editing one cell in a real report does not turn it into a bare table. The same applies to small
details inside a sheet: `sheetId` values referenced by revisions and extensions; the `ca` flag on
a formula that must recalculate every time; `cm`/`vm` indices linking a cell to rich-value metadata;
defined-name attributes such as comment and shortcut key.

```csharp
// Everything preserved by default
Workbook book = await Workbook.LoadAsync("report.xlsm");
book.Sheets[0]["B7"] = 42;
await book.SaveAsync("report-updated.xlsm");   // charts and macros intact

// Disable when the file is read-only and memory matters more
var options = new LoadOptions { PreserveUnknownParts = false };
```

How it works:

- **Parts stay put.** A worksheet is saved at the same path inside the package and with the same
  relationship id it arrived with. Relative references in preserved parts keep resolving, and ids
  referenced by raw markup (`pivotCaches`, `externalReferences`, `drawing`) are not reused for our
  own relationships.
- **All sheet positions count the same way.** A defined name’s `localSheetId` includes chart sheets
  and other non-modeled sheets. A workbook of chart sheets only does not gain an artificial
  `Sheet1` on save.
- **Reachability garbage collection.** Parts reachable from the root, workbook, and remaining
  sheets are copied, plus those that were tied to nothing in the source file. Parts from a deleted
  sheet are left behind instead of bloating the new file.
- **Worksheet markup by schema slot.** A worksheet is always regenerated, so its unrecognized
  elements return each to its own `CT_Worksheet` position — otherwise the file fails both validators
  and Excel.
- **Deliberately dropped.** `calcChain.xml` is not saved: our dependency graph is more reliable
  and Excel rebuilds the chain itself. The `dimension` element is not written — preserved from
  source it would be wrong after the first edit.
- **Broken relationships are repaired.** A relationship pointing at a missing part is not copied:
  such a file would not open. That is the only thing save changes on its own.

## Renaming a sheet

Renaming is not just the tab: formulas on every sheet, defined names (including print area),
validation and conditional-formatting formulas, internal hyperlinks, and in a loaded package chart
`c:f` references and pivot `worksheetSource` follow.

```csharp
book.Sheets["Data"].Name = "Q1 2026";
// SUM(Data!A1:A2) became SUM('Q1 2026'!A1:A2); quotes added per Excel rules
```

Formulas are edited as text, not through parsing: anything not recognized as a reference is copied
verbatim, so spaces, casing, and exotic syntax stay. A sheet name inside a string literal —
`"Data!A1"` — is not a reference and stays as-is. Saved charts and pivot sources are matched by
namespace URI and local name, not by whatever XML prefix the author chose (`c`, `x`, or any other).

### What is not covered

The full list of what load-save changes or loses — per the rule that any deviation from the spec
must be declared:

- **Built-in charts are read and created but not edited**: `sheet.Charts` shows type, title, and
  series (references + cached points); `sheet.AddChart(...)` creates a new one; raw `chartSpace` of
  an untouched chart is written back byte for byte. A loaded chart cannot be changed — only deleted
  and recreated. Chart sheets and pivot tables round-trip without API.
- **Sheet copy** does not take drawings and other raw markup from the source worksheet part, and
  does not rewrite formulas that name the source sheet by name.
- **Sheet rename** updates only chart `c:f` and pivot `worksheetSource`: if foreign markup hid a
  sheet name elsewhere, it stays old there.
- **Styles**: a `dxf` format the model cannot express (alignment, fill pattern) is parsed into
  `rule.Style` without those fields; the file keeps them until the rule style is rewritten. There
  is no API to create named cell styles — loaded ones are preserved.
- **Phonetic hints** (`rPh`, `phoneticPr`) survive load-save while cell runs are not rewritten, but
  are not modeled or regenerated: a hint points at a reading span that no longer exists after text
  replacement.
- **Data table formulas** (What-If, `f t="dataTable"`) are carried verbatim but never recalculated:
  the engine does not understand them. Load warns with code `DataTableFormula`.
- **`securityDescriptor`** on `protectedRange` is lost.
- **`calcChain.xml`** is not saved on purpose: our graph is more reliable; Excel will rebuild.
- **Strict files** (ISO conformance class) are read but do not survive save: the writer is always
  transitional and a package cannot be half one class and half the other. Load warns with code
  `StrictConformance` and disables unrecognized-part preservation, so the saved file contains only
  what the model knows.
- **`mc:MustUnderstand`** — extensions the file requires understanding are carried but not applied;
  load warns with code `MustUnderstand`.
- **`mc:AlternateContent`** is carried as original markup, but the reader does not choose or apply
  a `Choice` or `Fallback` branch. A value that exists only inside such a branch does not reach the
  model.
