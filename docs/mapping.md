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

- The type (and any containing types if it is nested) must be `partial` — the generator adds an
  `IRowSchema<T>` implementation (static abstract members).
- Positional records, record structs, and ordinary classes with public setters and a parameterless
  constructor work.
- `[ExcelColumn]`: header (defaults to the property name), `Order` (explicit position), `Format`
  (column number format on write).
- `[ExcelIgnore]` excludes a property.
- `[ExcelRow(HasHeader = false)]` — bind by position without a header row.

Supported types: `string`, `double`, `float`, `int`, `long`, `short`, `byte`, `decimal`, `bool`,
`DateTime`, `DateTimeOffset`, `DateOnly`, `TimeOnly`, `TimeSpan`, `Guid`, enums, `CellValue`, and
their nullable variants (blank ↔ `null`). An unsupported type is compile error `CWM007`.

## Reading

```csharp
// Streaming XLSX
await using XlsxReader reader = await XlsxReader.OpenAsync("orders.xlsx");
await foreach (Order order in reader.ReadRowsAsync<Order>("Orders"))
    Process(order);

// Streaming CSV/TSV
await foreach (Order order in CsvReader.ReadRowsAsync("orders.csv").ReadRowsAsync<Order>())
    Process(order);

// DOM Worksheet (from A1 or an explicit start cell)
IReadOnlyList<Order> domOrders = worksheet.ReadTable<Order>("B2");
```

The first row (for `HasHeader` schemas) maps columns by headers — case-insensitive and order-independent.
A missing header throws `CellMappingException` with a list of missing columns; a non-convertible cell
throws with the column name and value kind.

## Writing

```csharp
// Streaming XLSX: header + rows + column formats
await using XlsxWriter writer = await XlsxWriter.CreateAsync("orders.xlsx");
XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders");
await sheet.WriteRowsAsync(orders);

// Streaming CSV/TSV
await using CsvWriter csv = CsvWriter.Create("orders.csv");
await csv.WriteRowsAsync(orders);

// DOM: table from an arbitrary place on the sheet
worksheet.InsertTable(orders, "B2");
```

## Generator diagnostics

| Code | Situation |
| --- | --- |
| CWM001 | Type or one of its containing types is not `partial` |
| CWM002 | Reserved (nested `partial` types are supported) |
| CWM003 | Type is generic |
| CWM004 | No columns |
| CWM005 | Constructor parameter has no matching property |
| CWM006 | `[ExcelIgnore]` on a constructor parameter |
| CWM007 | Unsupported property type |
