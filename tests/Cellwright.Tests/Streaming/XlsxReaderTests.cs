using ClosedXML.Excel;
using Xunit;

namespace Cellwright.Tests.Streaming;

public sealed class XlsxReaderTests
{
    private static async Task<List<(int Row, CellValue[] Cells)>> ReadAllRowsAsync(XlsxReader reader, int sheetIndex = 0)
    {
        var rows = new List<(int, CellValue[])>();
        await foreach (RowData row in reader.ReadRowsAsync(sheetIndex, TestContext.Current.CancellationToken))
            rows.Add((row.RowNumber, row.ToArray()));
        return rows;
    }

    [Fact]
    public async Task RoundTrip_AllValueKinds_ThroughOwnWriter()
    {
        var timestamp = new DateTime(2026, 7, 24, 15, 30, 12);
        var elapsed = new TimeSpan(30, 15, 0);

        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow("text", 42.5, true, timestamp, elapsed, ExcelError.NotAvailable);
            sheet.WriteRow(CellValue.Blank, "after gap");
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Assert.Equal("Data", Assert.Single(reader.Sheets).Name);
        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);

        Assert.Equal(2, rows.Count);
        CellValue[] first = rows[0].Cells;
        Assert.Equal("text", first[0].GetText());
        Assert.Equal(42.5, first[1].GetNumber());
        Assert.True(first[2].GetBoolean());
        Assert.Equal(timestamp, first[3].GetDateTime());
        Assert.Equal(elapsed, first[4].GetTimeSpan());
        Assert.Equal(ExcelError.NotAvailable, first[5].GetError());

        Assert.Equal(2, rows[1].Row);
        Assert.True(rows[1].Cells[0].IsBlank);
        Assert.Equal("after gap", rows[1].Cells[1].GetText());
    }

    [Fact]
    public async Task ReadsWorkbookProducedByClosedXml()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet sheet = workbook.AddWorksheet("From CXML");
            sheet.Cell(1, 1).Value = "заголовок";
            sheet.Cell(1, 2).Value = 3.14159;
            sheet.Cell(2, 1).Value = new DateTime(2024, 12, 31, 23, 59, 0);
            sheet.Cell(2, 2).Value = true;
            sheet.Cell(3, 3).Value = XLError.CellReference;
            workbook.AddWorksheet("Hidden one").Hide();
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Assert.Equal(2, reader.Sheets.Count);
        Assert.False(reader.Sheets[0].IsHidden);
        Assert.True(reader.Sheets[1].IsHidden);

        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);
        Assert.Equal(3, rows.Count);
        Assert.Equal("заголовок", rows[0].Cells[0].GetText());
        Assert.Equal(3.14159, rows[0].Cells[1].GetNumber());
        Assert.Equal(new DateTime(2024, 12, 31, 23, 59, 0), rows[1].Cells[0].GetDateTime());
        Assert.True(rows[1].Cells[1].GetBoolean());
        Assert.Equal(ExcelError.Reference, rows[2].Cells[2].GetError());
    }

    [Fact]
    public async Task SparseSheet_PreservesRowNumbersAndGaps()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Sparse", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow(1);
            sheet.SkipRows(98);
            sheet.WriteRow(CellValue.Blank, CellValue.Blank, "C100");
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);

        Assert.Equal(2, rows.Count);
        Assert.Equal(1, rows[0].Row);
        Assert.Equal(100, rows[1].Row);
        Assert.Equal(3, rows[1].Cells.Length);
        Assert.True(rows[1].Cells[0].IsBlank);
        Assert.Equal("C100", rows[1].Cells[2].GetText());
    }

    [Fact]
    public async Task InlineStrings_AndDate1904_RoundTrip()
    {
        var date = new DateTime(2020, 6, 15, 8, 0, 0);
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream,
            new XlsxWriterOptions { LeaveOpen = true, UseInlineStrings = true, DateSystem = ExcelDateSystem.Date1904 },
            TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("S", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow("inline", date);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Assert.Equal(ExcelDateSystem.Date1904, reader.DateSystem);
        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);
        Assert.Equal("inline", rows[0].Cells[0].GetText());
        Assert.Equal(date, rows[0].Cells[1].GetDateTime());
    }

    [Fact]
    public async Task ControlCharacters_SurviveTheEscapeRoundTrip()
    {
        string tricky = "line1\nline2\tend\u0001 and literal _x0041_ text";
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("S", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow(tricky);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);
        Assert.Equal(tricky, rows[0].Cells[0].GetText());
    }

    [Fact]
    public async Task ReadRowsByName_IsCaseInsensitive_AndUnknownThrows()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Отчёт", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow(7);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        await foreach (RowData row in reader.ReadRowsAsync("ОТЧЁТ", TestContext.Current.CancellationToken))
            Assert.Equal(7.0, row.Cells[0].GetNumber());

        Assert.Throws<ArgumentException>(() => reader.ReadRowsAsync("нет такого", TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task MalformedContent_ProducesWarningsInsteadOfThrowing()
    {
        using var stream = new MemoryStream();
        BuildWorkbookWithBrokenSheet(stream);
        stream.Position = 0;

        var warnings = new List<WorkbookWarning>();
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream,
            new XlsxReaderOptions { LeaveOpen = true, OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        List<(int Row, CellValue[] Cells)> rows = await ReadAllRowsAsync(reader);

        Assert.Single(rows);
        Assert.Equal(5.0, rows[0].Cells[0].GetNumber());
        Assert.Contains(warnings, w => w.Code == "BadSharedStringIndex");
        Assert.Contains(warnings, w => w.Code == "BadNumber");
    }

    private static void BuildWorkbookWithBrokenSheet(Stream stream)
    {
        using var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true);
        WriteEntry(zip, "[Content_Types].xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);
        WriteEntry(zip, "_rels/.rels",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);
        WriteEntry(zip, "xl/_rels/workbook.xml.rels",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """);
        WriteEntry(zip, "xl/workbook.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets><sheet name="Broken" sheetId="1" r:id="rId1"/></sheets>
            </workbook>
            """);
        WriteEntry(zip, "xl/worksheets/sheet1.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData>
                <row r="1">
                  <c r="A1"><v>5</v></c>
                  <c r="B1" t="s"><v>999</v></c>
                  <c r="C1"><v>not-a-number</v></c>
                </row>
              </sheetData>
            </worksheet>
            """);
    }

    private static void WriteEntry(System.IO.Compression.ZipArchive zip, string name, string content)
    {
        System.IO.Compression.ZipArchiveEntry entry = zip.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }
}
