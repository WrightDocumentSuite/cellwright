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

public enum OrderPriority
{
    Low = 0,
    Normal = 1,
    High = 2,
}

public static partial class NestedRowContainer
{
    [ExcelRow]
    public sealed partial record ExtendedRow(
        Guid Id,
        OrderPriority Priority,
        OrderPriority? SecondaryPriority,
        float Ratio,
        short Rank,
        byte Flags,
        TimeOnly ScheduledAt,
        DateTimeOffset CreatedAt,
        [property: ExcelColumn("Keyword\nCol")] string @class);
}
