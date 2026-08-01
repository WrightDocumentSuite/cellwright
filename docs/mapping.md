# Typed Mapping (Cellwright.Mapping)

The `Cellwright.Mapping` package turns sheets into streams of typed rows and back. The whole
pipeline is generated at compile time by an incremental source generator — no reflection, fully
compatible with Native AOT and trimming.

## Declaring a schema

```csharp
using Cellwright.Mapping;

[ExcelRow]
public partial record Order(
    [ExcelColumn("Order ID")] int Id,
    string Customer,
    [ExcelColumn(Format = "#,##0.00")] decimal Total,
    DateTime? ShippedAt);
```

- The type must be `partial` and not nested — the generator adds an `IRowSchema<Order>`
  implementation (static abstract members).
- Positional records, record structs, and ordinary classes with public setters and a parameterless
  constructor work.
- `[ExcelColumn]`: header (defaults to the property name), `Order` (explicit position), `Format`
  (column number format on write).
- `[ExcelIgnore]` excludes a property.
- `[ExcelRow(HasHeader = false)]` — bind by position without a header row.

Supported types: `string`, `double`, `int`, `long`, `decimal`, `bool`, `DateTime`, `DateOnly`,
`TimeSpan`, `CellValue`, and their nullable variants (blank ↔ `null`). An unsupported type is compile
error `CWM007`.

## Reading

```csharp
await using XlsxReader reader = await XlsxReader.OpenAsync("orders.xlsx");
await foreach (Order order in reader.ReadRowsAsync<Order>("Orders"))
    Process(order);
```

The first row (for `HasHeader` schemas) maps columns by headers — case-insensitive and order-independent.
A missing header throws `CellMappingException` with a list of missing columns; a non-convertible cell
throws with the column name and value kind.

## Writing

```csharp
// Streaming: header + rows + column formats
await using XlsxWriter writer = await XlsxWriter.CreateAsync("orders.xlsx");
XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders");
await sheet.WriteRowsAsync(orders);

// DOM: table from an arbitrary place on the sheet
worksheet.InsertTable(orders, "B2");
```

## Generator diagnostics

| Code | Situation |
| --- | --- |
| CWM001 | Type is not `partial` |
| CWM002 | Type is nested inside another type |
| CWM003 | Type is generic |
| CWM004 | No columns |
| CWM005 | Constructor parameter has no matching property |
| CWM006 | `[ExcelIgnore]` on a constructor parameter |
| CWM007 | Unsupported property type |
