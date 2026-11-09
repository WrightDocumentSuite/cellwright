using System.Text;
using Xunit;

namespace Cellwright.Tests.Streaming;

/// <summary>
/// RFC 4180 CSV and TSV streaming reader, writer, and Workbook/Worksheet integration tests.
/// </summary>
public sealed class CsvTests
{
    private sealed class ForwardOnlyStream(byte[] content) : Stream
    {
        private readonly MemoryStream _inner = new(content);

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private static async Task<int> CountRowsAsync(Stream stream, CsvOptions options)
    {
        int rows = 0;
        await foreach (RowData row in CsvReader.ReadRowsAsync(stream, options, TestContext.Current.CancellationToken))
            rows++;
        return rows;
    }

    private static CsvOptions Limited(long bytes) =>
        new() { LoadBudget = DocumentLoadBudget.Default with { MaxInputBytes = bytes } };

    private static readonly byte[] TwentyByteRussian = Encoding.UTF8.GetBytes(new string('я', 10));

    [Fact]
    public async Task ByteLimitCountsBytesOnAForwardOnlyStream()
    {
        await Assert.ThrowsAsync<DocumentLoadLimitException>(
            () => CountRowsAsync(new ForwardOnlyStream(TwentyByteRussian), Limited(12)));
    }

    [Fact]
    public async Task ByteLimitGivesTheSameAnswerOnSeekableAndForwardOnlyStreams()
    {
        await Assert.ThrowsAsync<DocumentLoadLimitException>(
            () => CountRowsAsync(new MemoryStream(TwentyByteRussian), Limited(12)));

        Assert.Equal(1, await CountRowsAsync(new MemoryStream(TwentyByteRussian), Limited(20)));
        Assert.Equal(1, await CountRowsAsync(new ForwardOnlyStream(TwentyByteRussian), Limited(20)));
    }

    [Fact]
    public async Task ByteLimitSeesTheBytesOfUtf16()
    {
        byte[] content = [.. Encoding.Unicode.GetPreamble(), .. Encoding.Unicode.GetBytes("abcdefghij")];

        await Assert.ThrowsAsync<DocumentLoadLimitException>(
            () => CountRowsAsync(new ForwardOnlyStream(content), Limited(12)));
        Assert.Equal(1, await CountRowsAsync(new ForwardOnlyStream(content), Limited(22)));
    }

    [Fact]
    public async Task ByteLimitCountsFromTheCurrentPositionOfASeekableStream()
    {
        byte[] content = [.. new byte[100], .. Encoding.UTF8.GetBytes("abc")];
        var stream = new MemoryStream(content) { Position = 100 };

        Assert.Equal(1, await CountRowsAsync(stream, Limited(3)));
    }

    [Fact]
    public void ACharacterLimitStillAppliesToATextReader()
    {
        using var reader = new StringReader("abcdefghij");

        Assert.Throws<DocumentLoadLimitException>(() => CsvReader.ReadRows(reader, Limited(5)).ToList());
    }

    [Fact]
    public async Task TheStreamIsLeftOpenWhenAsked()
    {
        var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b"));

        await CountRowsAsync(stream, new CsvOptions { LeaveOpen = true });

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void TextAfterAClosingQuoteIsKeptByDefault()
    {
        Workbook workbook = Workbook.ParseCsv("\"good\"BAD,next");
        Worksheet sheet = workbook.Sheets["Sheet1"];

        Assert.Equal("goodBAD", sheet["A1"].GetText());
        Assert.Equal("next", sheet["B1"].GetText());
    }

    [Fact]
    public void StrictQuotesRefuseTextAfterAClosingQuote()
    {
        var options = new CsvOptions { StrictQuotes = true };

        Assert.Throws<InvalidDataException>(() => Workbook.ParseCsv("\"good\"BAD,next", options));
    }

    [Fact]
    public void StrictQuotesRefuseAnUnclosedQuote()
    {
        var options = new CsvOptions { StrictQuotes = true };

        Assert.Throws<InvalidDataException>(() => Workbook.ParseCsv("a,\"never closed", options));
        Assert.Equal("never closed", Workbook.ParseCsv("a,\"never closed").Sheets["Sheet1"]["B1"].GetText());
    }

    [Fact]
    public void StrictQuotesAcceptWellFormedFields()
    {
        var options = new CsvOptions { StrictQuotes = true };

        Worksheet sheet = Workbook.ParseCsv("\"a,b\",\"c\"\"d\"\r\n", options).Sheets["Sheet1"];

        Assert.Equal("a,b", sheet["A1"].GetText());
        Assert.Equal("c\"d", sheet["B1"].GetText());
    }

    [Fact]
    public void ParseCsvAndToCsvRoundTripTypesQuotesAndLeadingZeros()
    {
        string csv =
            "Name,Amount,Active,Date,Duration,Code,Notes\r\n" +
            "\"Alice, Jr.\",1250.5,TRUE,2026-10-16,01:30:45,00123,\"Line 1\n\"\"Quoted\"\"\"\r\n" +
            "Bob,-42,FALSE,2026-01-01,-00:15:00,\"00456\",\r\n";

        Workbook workbook = Workbook.ParseCsv(csv, new CsvOptions { SheetName = "Imports" });
        Worksheet sheet = workbook.Sheets["Imports"];

        Assert.Equal("Alice, Jr.", sheet["A2"].GetText());
        Assert.Equal(1250.5, sheet["B2"].GetNumber());
        Assert.True(sheet["C2"].GetBoolean());
        Assert.Equal(new DateTime(2026, 10, 16), sheet["D2"].GetDateTime());
        Assert.Equal(new TimeSpan(1, 30, 45), sheet["E2"].GetTimeSpan());
        Assert.Equal("00123", sheet["F2"].GetText());
        Assert.Equal("Line 1\n\"Quoted\"", sheet["G2"].GetText());

        Assert.Equal(new TimeSpan(0, -15, 0), sheet["E3"].GetTimeSpan());
        Assert.Equal("00456", sheet["F3"].GetText());
        Assert.True(sheet["G3"].IsBlank);

        string exported = sheet.ToCsv();
        Workbook roundTripped = Workbook.ParseCsv(exported);
        Worksheet rtSheet = roundTripped.Sheets[0];

        Assert.Equal("Alice, Jr.", rtSheet["A2"].GetText());
        Assert.Equal(1250.5, rtSheet["B2"].GetNumber());
        Assert.True(rtSheet["C2"].GetBoolean());
        Assert.Equal(new DateTime(2026, 10, 16), rtSheet["D2"].GetDateTime());
        Assert.Equal(new TimeSpan(1, 30, 45), rtSheet["E2"].GetTimeSpan());
        Assert.Equal("00123", rtSheet["F2"].GetText());
        Assert.Equal("Line 1\n\"Quoted\"", rtSheet["G2"].GetText());
    }

    [Fact]
    public async Task StreamingCsvReaderAndWriterSupportTabSeparatedValues()
    {
        using var stream = new MemoryStream();
        CsvOptions options = CsvOptions.TabSeparated with { LeaveOpen = true };

        await using (CsvWriter writer = CsvWriter.Create(stream, options))
        {
            await writer.WriteRowAsync(
                new[] { CellValue.Text("Region"), CellValue.Text("Total") },
                TestContext.Current.CancellationToken);
            await writer.WriteRowAsync(
                new[] { CellValue.Text("North\tEast"), CellValue.Number(300) },
                TestContext.Current.CancellationToken);
        }

        stream.Position = 0;
        var rows = new List<CellValue[]>();
        await foreach (RowData row in CsvReader.ReadRowsAsync(stream, options, TestContext.Current.CancellationToken))
            rows.Add(row.ToArray());

        Assert.Equal(2, rows.Count);
        Assert.Equal("Region", rows[0][0].GetText());
        Assert.Equal("North\tEast", rows[1][0].GetText());
        Assert.Equal(300d, rows[1][1].GetNumber());
    }

    [Fact]
    public async Task LoadCsvAsyncAndSaveCsvAsyncWorkWithStreams()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Item;Price\r\nBook;19.99\r\n"));
        CsvOptions options = CsvOptions.SemicolonSeparated with { LeaveOpen = true };

        Workbook workbook = await Workbook.LoadCsvAsync(stream, options, TestContext.Current.CancellationToken);
        Assert.Equal("Book", workbook.Sheets[0]["A2"].GetText());
        Assert.Equal(19.99, workbook.Sheets[0]["B2"].GetNumber());

        using var output = new MemoryStream();
        await workbook.Sheets[0].SaveCsvAsync(output, options, TestContext.Current.CancellationToken);
        string written = Encoding.UTF8.GetString(output.ToArray());
        Assert.Equal("Item;Price\r\nBook;19.99\r\n", written);
    }

    [Fact]
    public async Task StreamingCsvReaderHandlesRecordsCrossingChunkBoundariesAndEnforcesLoadBudget()
    {
        var sb = new StringBuilder();
        for (int i = 1; i <= 500; i++)
            sb.Append($"\"Row {i} with \"\"escaped\"\" text\",{i},02:15:00\r\n");

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));
        int rowCount = 0;
        await foreach (RowData row in CsvReader.ReadRowsAsync(
            stream, new CsvOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            rowCount++;
            Assert.Equal(rowCount, row.RowNumber);
            Assert.Equal((double)rowCount, row.Cells[1].GetNumber());
            Assert.Equal(new TimeSpan(2, 15, 0), row.Cells[2].GetTimeSpan());
        }

        Assert.Equal(500, rowCount);

        stream.Position = 0;
        var tightOptions = new CsvOptions
        {
            LeaveOpen = true,
            LoadBudget = new DocumentLoadBudget { MaxCells = 10 },
        };

        await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
        {
            await foreach (RowData _ in CsvReader.ReadRowsAsync(stream, tightOptions, TestContext.Current.CancellationToken))
            {
            }
        });
    }
}
