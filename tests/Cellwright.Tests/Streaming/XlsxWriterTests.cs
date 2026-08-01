using System.IO.Compression;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Streaming;

public sealed class XlsxWriterTests
{
    private static async Task<MemoryStream> WriteWorkbookAsync(Func<XlsxWriter, ValueTask> populate, XlsxWriterOptions? options = null)
    {
        var stream = new MemoryStream();
        XlsxWriterOptions effective = (options ?? new XlsxWriterOptions()) with { LeaveOpen = true };
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(stream, effective, TestContext.Current.CancellationToken))
        {
            await populate(writer);
        }

        stream.Position = 0;
        return stream;
    }

    private static void AssertPackageIsValid(MemoryStream stream)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document, TestContext.Current.CancellationToken).ToList();
        Assert.Empty(errors);
        stream.Position = 0;
    }

    [Fact]
    public async Task AllValueKinds_ValidateAndRoundTripThroughClosedXml()
    {
        var timestamp = new DateTime(2026, 7, 24, 15, 30, 0);
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data");
            sheet.WriteRow("Name", "Score", "Passed", "When", "Elapsed", "Error");
            sheet.WriteRow("Анна", 99.5, true, timestamp, new TimeSpan(30, 15, 0), ExcelError.DivisionByZero);
            sheet.WriteRow(CellValue.Blank, -0.25, false);
        });

        AssertPackageIsValid(stream);

        using var workbook = new XLWorkbook(stream);
        IXLWorksheet sheet = workbook.Worksheet("Data");

        Assert.Equal("Name", sheet.Cell(1, 1).GetString());
        Assert.Equal("Анна", sheet.Cell(2, 1).GetString());
        Assert.Equal(99.5, sheet.Cell(2, 2).GetDouble());
        Assert.True(sheet.Cell(2, 3).GetBoolean());
        Assert.Equal(timestamp, sheet.Cell(2, 4).GetDateTime());
        Assert.Equal(new TimeSpan(30, 15, 0), sheet.Cell(2, 5).GetTimeSpan());
        Assert.Equal(XLError.DivisionByZero, sheet.Cell(2, 6).GetError());
        Assert.True(sheet.Cell(3, 1).Value.IsBlank);
        Assert.Equal(-0.25, sheet.Cell(3, 2).GetDouble());
    }

    [Fact]
    public async Task SharedStrings_AreDeduplicated()
    {
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("S");
            for (int i = 0; i < 50; i++)
                sheet.WriteRow("repeated", "repeated", "unique-" + (i % 2));
        });

        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry? entry = zip.GetEntry("xl/sharedStrings.xml");
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        string xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
        Assert.Contains("uniqueCount=\"3\"", xml, StringComparison.Ordinal);
        Assert.Contains("count=\"150\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task InlineStringsMode_SkipsSharedStringsPart()
    {
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("S");
            sheet.WriteRow("inline text", "  padded  ");
        }, new XlsxWriterOptions { UseInlineStrings = true });

        AssertPackageIsValid(stream);

        using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
        {
            Assert.Null(zip.GetEntry("xl/sharedStrings.xml"));
        }

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        Assert.Equal("inline text", workbook.Worksheet("S").Cell(1, 1).GetString());
        Assert.Equal("  padded  ", workbook.Worksheet("S").Cell(1, 2).GetString());
    }

    [Fact]
    public async Task EmptyWorkbook_GetsDefaultSheet()
    {
        using MemoryStream stream = await WriteWorkbookAsync(_ => ValueTask.CompletedTask);
        AssertPackageIsValid(stream);

        using var workbook = new XLWorkbook(stream);
        Assert.Equal("Sheet1", Assert.Single(workbook.Worksheets).Name);
    }

    [Fact]
    public async Task MultipleSheets_PreviousSheetAutoCompletes()
    {
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter first = await writer.BeginSheetAsync("First");
            first.WriteRow(1);
            XlsxSheetWriter second = await writer.BeginSheetAsync("Second");
            second.WriteRow(2);
        });

        AssertPackageIsValid(stream);
        using var workbook = new XLWorkbook(stream);
        Assert.Equal(1, workbook.Worksheet("First").Cell(1, 1).GetDouble());
        Assert.Equal(2, workbook.Worksheet("Second").Cell(1, 1).GetDouble());
    }

    [Fact]
    public async Task MergesFrozenPanesAndWidths_RoundTrip()
    {
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Report", new SheetOptions
            {
                FrozenRows = 1,
                FrozenColumns = 2,
                Columns = new Dictionary<int, ColumnProperties>
                {
                    [1] = new() { Width = 25.5 },
                    [3] = new() { Width = 10 },
                },
            });
            sheet.WriteRow("Header spanning");
            sheet.Merge(RangeRef.Parse("A1:C1"));
            sheet.SkipRows(2);
            sheet.WriteRow(CellValue.Blank, "B4 value");
        });

        AssertPackageIsValid(stream);

        using var workbook = new XLWorkbook(stream);
        IXLWorksheet sheet = workbook.Worksheet("Report");
        Assert.Equal("A1:C1", Assert.Single(sheet.MergedRanges).RangeAddress.ToString());
        Assert.Equal(1, sheet.SheetView.SplitRow);
        Assert.Equal(2, sheet.SheetView.SplitColumn);
        // ClosedXML converts the raw OOXML width into its own character units (~0.71 smaller).
        Assert.InRange(sheet.Column(1).Width, 24.5, 25.5);
        Assert.Equal("B4 value", sheet.Cell(4, 2).GetString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("name:with:colons")]
    [InlineData("way/too?bad*[name]")]
    [InlineData("'quoted'")]
    [InlineData("123456789012345678901234567890123")]
    public async Task InvalidSheetNames_Throw(string name)
    {
        using var stream = new MemoryStream();
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        await Assert.ThrowsAnyAsync<ArgumentException>(async () =>
            await writer.BeginSheetAsync(name, cancellationToken: TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task DuplicateSheetName_Throws()
    {
        using var stream = new MemoryStream();
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        await writer.BeginSheetAsync("Same", cancellationToken: TestContext.Current.CancellationToken);
        await Assert.ThrowsAsync<ArgumentException>(async () =>
            await writer.BeginSheetAsync("same", cancellationToken: TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task LargeSheet_StreamsWithPeriodicFlushes()
    {
        using MemoryStream stream = await WriteWorkbookAsync(async writer =>
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Big");
            for (int i = 1; i <= 5_000; i++)
            {
                sheet.WriteRow(i, "row-" + (i % 100), i * 0.5, i % 2 == 0);
                await sheet.FlushIfNeededAsync(TestContext.Current.CancellationToken);
            }
        });

        using var workbook = new XLWorkbook(stream);
        IXLWorksheet sheet = workbook.Worksheet("Big");
        Assert.Equal(5_000, sheet.LastRowUsed()!.RowNumber());
        Assert.Equal(2_500.0, sheet.Cell(5_000, 3).GetDouble());
    }
}
