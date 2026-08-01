# Styles

## Model

A style is an immutable `CellStyle` record with six parts:

```csharp
var style = new CellStyle
{
    Font = new FontStyle { Bold = true, Size = 13, Color = ExcelColor.FromRgb(255, 255, 255) },
    Fill = FillStyle.Solid(ExcelColor.FromArgb(0xFF2278D4)),
    Border = BorderStyle.All(BorderEdge.Thin),
    Alignment = new Alignment { Horizontal = HorizontalAlignment.Center, WrapText = true },
    Protection = CellProtection.Unlocked,
    NumberFormat = "#,##0.00",
};
```

Records compare by value, so equal styles automatically collapse to one format on write — register
them as often as you like.

Coverage per ISO/IEC 29500 §18.8:

- **Font**: name, size, bold, italic, strikethrough, five underline kinds (`UnderlineStyle`),
  superscript and subscript (`VerticalAlign`), color, `Family`, `Charset`, and `Scheme` — the
  theme slot that makes the font follow the workbook theme.
- **Fill**: solid and all 17 pattern fills (`FillPattern`) with foreground and background colors,
  plus gradients (§18.8.24) — linear with angle and radial (`path`) with convergence inserts;
  stops with any number of colors: `FillStyle.Gradiented(GradientFill.Linear(from, to, 90))`.
- **Borders**: four sides and diagonal (`Diagonal` + `DiagonalUp`/`DiagonalDown`), all 13
  `ST_BorderStyle` line kinds.
- **Alignment**: horizontal and vertical in every schema variant, wrap, indent, text rotation
  (`TextRotation`, 255 — stacked letters), `ShrinkToFit`, reading direction.
- **Protection**: `Locked` and `Hidden` — by default a cell is locked, as in Excel; meaning
  appears only after `sheet.Protection.Protect()`.

## Colors

```csharp
ExcelColor.FromRgb(0x22, 0x78, 0xD4);   // explicit ARGB
ExcelColor.FromArgb(0xFF2278D4);
ExcelColor.FromTheme(4, tint: -0.25);   // theme slot, darkened
ExcelColor.Auto;                        // consumer decides
```

Excel’s default palette is theme-based, so a theme color is not exotic — it is how most files are
colored. The theme part is carried into the saved file as-is, so the slot stays meaningful even
though the library does not resolve it to pixels. Legacy `indexed` colors resolve on read through a
fixed 64-entry palette.

## Fluent builders (C# 14 extension members)

The type stays a pure record; builders are extension members:

```csharp
CellStyle header = CellStyle.Default.Bold().FontSize(14).Fill(0xFF2278D4).Center().Wrap();
```

Available: `Bold`, `Italic`, `Underline`, `Strikethrough`, `FontName`, `FontSize`, `FontColor`,
`Fill`, `Border`, `NumberFormat`, `Align`, `Center`, `Wrap`, `Unlocked`, `HideFormula`, and
`IsDefault`.

## Applying styles

DOM — through facades (each mutator returns the facade for chaining):

```csharp
StyleId id = book.RegisterStyle(header);      // handle; equal styles get equal handles
sheet.Cell("A1").SetStyle(id);
sheet.Cell("B1").SetStyle(header);            // or pass the record — registered automatically
sheet.Cell("C1").Bold();                      // shortcuts read-modify-intern
sheet.Range("A2:D9").Style(s => s.NumberFormat("0.00%"));  // transform, memoized from source style
```

Streaming — a parallel span of styles:

```csharp
StyleId money = writer.RegisterStyle(CellStyle.Default.NumberFormat("#,##0.00"));
sheet.WriteRow([name, total], [StyleId.Default, money]);
```

A row and column carry their own style — cells without one inherit it. One record formats the
whole row, not sixteen thousand cells:

```csharp
sheet.Row(1).SetStyle(header);
sheet.Column("D").SetStyle(CellStyle.Default.NumberFormat("0.00%"));
```

## Number formats

`CellStyle.NumberFormat` is a format code (`""` = General). Codes that match a builtin (for example
`0.00%`, `#,##0.00`, `m/d/yy h:mm`) are written with a builtin id; others get custom ids (164+).
On read, date-like formats automatically turn serial numbers into `DateTime`/`TimeSpan`.

## Named cell styles

Masters in `cellStyleXfs` and the name table `cellStyles` (“Normal”, “Heading 1”, currency presets)
survive load and save: masters parse into the same `CellStyle` records — their component indices
point into tables the writer rebuilds, so byte copying here would lie — and each cell format’s
`xfId` is preserved and still names its master. A style that differs only by named-style membership
does not collapse with a duplicate: record equality deliberately includes `xfId`. There is no API
to create named styles — a workbook built from scratch carries one `Normal`.

## Rich text in a cell

A cell style sets one font for the whole cell. To format part of the text, the format stores a
sequence of runs (`r`), each with its own font (`rPr`). That is `RichText`:

```csharp
FontStyle body = sheet.Cell("A1").StyleValue.Font;

sheet.Cell("A1").SetRichText(RichText.Empty
    .Append("Total: ")
    .Append("42", body with { Bold = true }));

RichText? runs = sheet.Cell("A1").RichText;   // null if no rich formatting
string visible = sheet.Cell("A1").Value.GetText();  // "Total: 42"
```

A run without a font is drawn with the cell font — a run without `rPr`. A run that names a font
names it **whole**: the format cannot say “like the cell but bold”, so a bold run carries family,
size, and color too. Hence the pattern above — take the cell font and change one field; otherwise
the run silently gets Calibri 11 instead of the cell font.

Formatting belongs to the text it decorates. Replacing the cell value removes it — `SetValue`,
`Clear`, and indexer assignment — and `RichText` then returns `null`, not word-level markup for
words that no longer exist. Text whose runs all follow the cell font is stored as a plain string:
the runs would have nothing to say.

The same visible text with different run sets stays different shared-string entries — the format
allows that, and collapsing them would lose formatting.

Not covered here: phonetic hints (`rPh`, `phoneticPr`) are not modeled — see
[worksheet.md](worksheet.md). In `.xls` and `.xlsb`, runs collapse to plain text.

## What is not covered

Differential formats (`dxf`) from a loaded file parse into the model — a conditional-formatting
rule exposes them through `Style` — but the format table is still written from the original markup,
and a format the model cannot express (alignment in a dxf, fill pattern) loses those fields in the
parsed view; the file is unaffected until the rule style is rewritten. Gradients inside a `dxf` are
not parsed.
