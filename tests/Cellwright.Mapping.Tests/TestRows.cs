using Cellwright.Mapping;

namespace Cellwright.Mapping.Tests;

[ExcelRow]
public partial record Order(
    [ExcelColumn("Order ID")] int Id,
    string Customer,
    [ExcelColumn(Format = "#,##0.00")] decimal Total,
    DateTime? ShippedAt);

[ExcelRow(HasHeader = false)]
public partial class PlainRow
{
    public string? Name { get; set; }

    public double Score { get; set; }

    [ExcelIgnore]
    public string? Secret { get; set; }
}

[ExcelRow]
public partial record struct Measurement(
    [ExcelColumn("Метка времени")] DateTime Timestamp,
    double? Value,
    bool Valid);
