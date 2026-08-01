# Getting Started

Requires the .NET 10 SDK. Packages: `Cellwright` (core) and optionally `Cellwright.Mapping`
(typed mapping with a source generator).

## Create and save a workbook (DOM)

```csharp
using Cellwright;

var book = Workbook.Create();
Worksheet sheet = book.Sheets.Add("Sales");

sheet["A1"] = "Product";                  // indexer — shortest path to a value
sheet["B1"] = "Amount";
sheet[2, 1] = "Apples";                   // (row, column), 1-based
sheet[2, 2] = 1234.5;
sheet["A3"] = DateTime.Today;             // dates get a format automatically
sheet.Cell("B3").SetFormula("SUM(B2:B2)"); // formulas as text; Excel recalculates on open

await book.SaveAsync("sales.xlsx");
```

## Load, edit, save

```csharp
Workbook book = await Workbook.LoadAsync("sales.xlsx", new LoadOptions
{
    OnWarning = w => Console.WriteLine(w),   // diagnostics for recoverable issues
});

Worksheet sheet = book.Sheets["Sales"];
Console.WriteLine(sheet["B2"].GetNumber());
sheet["C1"] = "Updated";

await book.SaveAsync("sales-updated.xlsx");
```

`CellValue` is a compact tagged union: inspect the kind with `IsNumber`/`IsText`/... and read
with `GetNumber()`/`GetText()`/`TryGetNumber(...)`. Implicit conversions exist for all
primitives, including nullable (null becomes `Blank`).

## Styles

```csharp
sheet.Range("A1:B1").Style(s => s.Bold().Fill(0xFF2278D4).FontColor(ExcelColor.FromRgb(255, 255, 255)));
sheet.Range("B2:B9").NumberFormat("#,##0.00");
sheet.Cell("A5").SetValue("Total").Bold();
```

## Rows, columns, and other worksheet features

```csharp
sheet.Row(1).SetHeight(24);
sheet.Column("A").SetWidth(24);
sheet.Rows(5, 12).Group();
sheet.AutoFilter("A1:B9");
sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com");
sheet.Comments.Add("B2", "Check the number", author: "Auditor");
sheet.Validations.Add(DataValidation.List("C2:C99", "\"Yes,No\""));
sheet.Tables.Add("Sales", "A1:B9");

book.Sheets.Copy("Sales", "Sales (copy)");   // copy is inserted right after the original
book.Sheets.Move("Sales", 0);
book.Properties.Author = "Anna Smith";       // Excel Info tab
```

Editing a loaded file does not destroy anything: charts, pivot tables, pictures, and macros
survive in the saved file as they were. See [worksheet.md](worksheet.md) for details.

## Unknown file format

`Cellwright.Xls` opens both legacy `.xls` and password-protected packages — you cannot trust the
extension or the first bytes alone:

```csharp
using Cellwright.Xls;

Workbook book = await Spreadsheet.LoadAsync(uploaded, new XlsLoadOptions { Password = "secret" });
```

Without that package the core does not fail silently: wrong-format files — `.xlsb`, OpenDocument,
encrypted packages — are named explicitly in `XlsxFormatException`.

## Formulas

The `Cellwright.Formulas` package adds calculation:

```csharp
using Cellwright.Formulas;

sheet.Cell("B3").SetFormula("SUM(B1:B2)*1.2");
book.Calculate();                       // first pass calculates everything
sheet["B1"] = 500;
book.Calculate();                       // later passes only dependents
Console.WriteLine(sheet["B3"].GetNumber());
```

See [formulas.md](formulas.md) for calculation details. Streaming millions of rows —
[streaming.md](streaming.md). Worksheet features — [worksheet.md](worksheet.md). Typed rows —
[mapping.md](mapping.md). Styles — [styles.md](styles.md).
