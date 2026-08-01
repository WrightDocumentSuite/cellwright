using Xunit;

namespace Cellwright.Mapping.Tests;

public sealed class MappingRoundTripTests
{
    private static readonly List<Order> SampleOrders =
    [
        new(1, "Анна", 1234.56m, new DateTime(2026, 7, 1)),
        new(2, "Boris", 99.99m, null),
        new(3, "Clara", 0.5m, new DateTime(2026, 7, 20, 14, 30, 0)),
    ];

    [Fact]
    public async Task StreamingWriteRead_TypedRows()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders", cancellationToken: TestContext.Current.CancellationToken);
            await sheet.WriteRowsAsync(SampleOrders, TestContext.Current.CancellationToken);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        var roundTripped = new List<Order>();
        await foreach (Order order in reader.ReadRowsAsync<Order>("Orders", TestContext.Current.CancellationToken))
            roundTripped.Add(order);

        Assert.Equal(SampleOrders, roundTripped);
    }

    [Fact]
    public async Task HeaderBinding_SurvivesReorderedColumns()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders", cancellationToken: TestContext.Current.CancellationToken);
            // Columns deliberately shuffled relative to the schema declaration.
            sheet.WriteRow("Customer", "Total", "Order ID", "ShippedAt");
            sheet.WriteRow("Дарья", 55.5, 7, new DateTime(2026, 1, 2));
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Order order = await Single(reader.ReadRowsAsync<Order>("Orders", TestContext.Current.CancellationToken));
        Assert.Equal(new Order(7, "Дарья", 55.5m, new DateTime(2026, 1, 2)), order);
    }

    [Fact]
    public async Task MissingHeader_ThrowsMappingException()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Orders", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow("Order ID", "Customer"); // Total and ShippedAt are missing.
            sheet.WriteRow(1, "x");
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        CellMappingException exception = await Assert.ThrowsAsync<CellMappingException>(async () =>
        {
            await foreach (Order _ in reader.ReadRowsAsync<Order>("Orders", TestContext.Current.CancellationToken))
            {
            }
        });
        Assert.Contains("Total", exception.Message, StringComparison.Ordinal);
        Assert.Contains("ShippedAt", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task HeaderlessPoco_BindsPositionally_AndIgnoresExcludedProperty()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data", cancellationToken: TestContext.Current.CancellationToken);
            await sheet.WriteRowsAsync(
                [new PlainRow { Name = "first", Score = 1.5, Secret = "never" }],
                TestContext.Current.CancellationToken);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        PlainRow row = await Single(reader.ReadRowsAsync<PlainRow>("Data", TestContext.Current.CancellationToken));
        Assert.Equal("first", row.Name);
        Assert.Equal(1.5, row.Score);
        Assert.Null(row.Secret);
    }

    [Fact]
    public async Task RecordStruct_WithNullableAndUnicodeHeader_RoundTrips()
    {
        var source = new Measurement(new DateTime(2026, 7, 24, 10, 0, 0), null, true);

        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("M", cancellationToken: TestContext.Current.CancellationToken);
            await sheet.WriteRowsAsync([source], TestContext.Current.CancellationToken);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Measurement result = await Single(reader.ReadRowsAsync<Measurement>("M", TestContext.Current.CancellationToken));
        Assert.Equal(source, result);
    }

    [Fact]
    public async Task InsertTable_WritesHeaderRowsAndStyles()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Orders");
        sheet.InsertTable(SampleOrders, "B2");

        Assert.Equal("Order ID", sheet["B2"].GetText());
        Assert.Equal("Customer", sheet["C2"].GetText());
        Assert.Equal(1.0, sheet["B3"].GetNumber());
        Assert.Equal("Анна", sheet["C3"].GetText());
        Assert.Equal("#,##0.00", sheet.Cell("D3").StyleValue.NumberFormat);
        Assert.True(sheet["E4"].IsBlank); // Null ShippedAt.

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.True(stream.Length > 0);
    }

    private static async Task<T> Single<T>(IAsyncEnumerable<T> source)
    {
        T? result = default;
        int count = 0;
        await foreach (T item in source)
        {
            result = item;
            count++;
        }

        Assert.Equal(1, count);
        return result!;
    }
}
