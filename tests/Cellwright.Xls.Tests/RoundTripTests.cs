using System.Text;
using ExcelDataReader;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Writes a workbook in the legacy format and reads it back — with this library, and with an
/// independent one. The second half is what makes the test worth anything: a writer and a
/// reader that share a misunderstanding will agree with each other perfectly.
/// </summary>
public sealed class RoundTripTests
{
    static RoundTripTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task ValuesSurviveAWriteAndRead()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Data");
        sheet["A1"] = CellValue.Text("Name");
        sheet["B1"] = CellValue.Text("Amount");
        sheet["A2"] = CellValue.Text("Widget");
        sheet["B2"] = CellValue.Number(12.5);
        sheet["A3"] = CellValue.Text("Gadget");
        sheet["B3"] = CellValue.Number(1234);
        sheet["A4"] = CellValue.Boolean(true);
        sheet["B4"] = CellValue.Error(ExcelError.DivisionByZero);
        sheet["A5"] = CellValue.Number(-0.25);

        Workbook reloaded = await RoundTripAsync(original);

        Worksheet read = reloaded.Sheets[0];
        Assert.Equal("Data", read.Name);
        Assert.Equal("Name", read["A1"].GetText());
        Assert.Equal("Widget", read["A2"].GetText());
        Assert.Equal(12.5, read["B2"].GetNumber());
        Assert.Equal(1234, read["B3"].GetNumber());
        Assert.True(read["A4"].GetBoolean());
        Assert.Equal(ExcelError.DivisionByZero, read["B4"].GetError());
        Assert.Equal(-0.25, read["A5"].GetNumber());
    }

    [Fact]
    public async Task SeveralSheetsKeepTheirNamesAndOrder()
    {
        Workbook original = Workbook.Create();
        foreach (string name in new[] { "First", "Second", "Third" })
        {
            Worksheet sheet = original.Sheets.Add(name);
            sheet["A1"] = CellValue.Text(name);
        }

        original.Sheets[1].Visibility = SheetVisibility.Hidden;

        Workbook reloaded = await RoundTripAsync(original);

        Assert.Equal(3, reloaded.Sheets.Count);
        Assert.Equal(["First", "Second", "Third"], reloaded.Sheets.Select(static s => s.Name));
        Assert.Equal("Third", reloaded.Sheets[2]["A1"].GetText());
        Assert.Equal(SheetVisibility.Hidden, reloaded.Sheets[1].Visibility);
    }

    [Fact]
    public async Task FormulasKeepTheirTextAndTheirLastValue()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Calc");
        sheet["A1"] = CellValue.Number(2);
        sheet["A2"] = CellValue.Number(3);
        sheet.Cell("A3").SetFormula("SUM(A1:A2)");
        sheet.Values.Set(3, 1, CellValue.Number(5));
        sheet.Cell("B1").SetFormula("IF(A1>1,\"big\",\"small\")");
        sheet.Values.Set(1, 2, CellValue.Text("big"));

        Workbook reloaded = await RoundTripAsync(original);

        Worksheet read = reloaded.Sheets[0];
        Assert.Equal("SUM(A1:A2)", read.Cell("A3").Formula);
        Assert.Equal(5, read["A3"].GetNumber());
        Assert.Equal("IF(A1>1,\"big\",\"small\")", read.Cell("B1").Formula);
        Assert.Equal("big", read["B1"].GetText());
    }

    [Fact]
    public async Task DatesComeBackAsDates()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Dates");
        StyleId dateStyle = original.RegisterStyle(CellStyle.Default with { NumberFormat = "yyyy-mm-dd" });
        sheet["A1"] = CellValue.DateTime(new DateTime(2024, 3, 17, 0, 0, 0, DateTimeKind.Unspecified));
        sheet.Cell("A1").SetStyle(dateStyle);

        Workbook reloaded = await RoundTripAsync(original);

        Assert.Equal(new DateTime(2024, 3, 17, 0, 0, 0, DateTimeKind.Unspecified), reloaded.Sheets[0]["A1"].GetDateTime());
    }

    [Fact]
    public async Task StylesSurviveTheirOwnTable()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Styled");
        StyleId bold = original.RegisterStyle(CellStyle.Default with
        {
            Font = FontStyle.Default with { Bold = true, Size = 14, Name = "Times New Roman" },
            Fill = FillStyle.Solid(ExcelColor.FromRgb(0xFF, 0xC0, 0x00)),
            Alignment = Alignment.Default with { Horizontal = HorizontalAlignment.Center, WrapText = true },
            NumberFormat = "#,##0.00",
        });

        sheet["A1"] = CellValue.Number(1234.5);
        sheet.Cell("A1").SetStyle(bold);

        Workbook reloaded = await RoundTripAsync(original);

        CellStyle style = reloaded.GetStyle(reloaded.Sheets[0].Cell("A1").Style);
        Assert.True(style.Font.Bold);
        Assert.Equal(14, style.Font.Size);
        Assert.Equal("Times New Roman", style.Font.Name);
        Assert.Equal("#,##0.00", style.NumberFormat);
        Assert.Equal(HorizontalAlignment.Center, style.Alignment.Horizontal);
        Assert.True(style.Alignment.WrapText);
        Assert.Equal(FillPattern.Solid, style.Fill.Pattern);
        Assert.Equal(0xFFFFC000u, style.Fill.Color.Argb);
    }

    [Fact]
    public async Task MergedRangesAndFrozenPanesSurvive()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Layout");
        sheet["A1"] = CellValue.Text("Title");
        sheet.Merge(RangeRef.Parse("A1:C1"));
        sheet.FrozenRows = 1;
        sheet.FrozenColumns = 2;
        sheet.ShowGridLines = false;

        Workbook reloaded = await RoundTripAsync(original);

        Worksheet read = reloaded.Sheets[0];
        Assert.Equal(RangeRef.Parse("A1:C1"), Assert.Single(read.MergedRanges));
        Assert.Equal(1, read.FrozenRows);
        Assert.Equal(2, read.FrozenColumns);
        Assert.False(read.ShowGridLines);
    }

    /// <summary>
    /// A long shared-string table has to be split across records, and the split is the part of
    /// the format most easily got wrong. Two thousand distinct strings force several.
    /// </summary>
    [Fact]
    public async Task ALargeStringTableSurvivesBeingSplit()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Strings");
        for (int row = 1; row <= 2000; row++)
            sheet[row, 1] = CellValue.Text($"Row {row} of a table long enough to need continuation records");

        Workbook reloaded = await RoundTripAsync(original);

        Worksheet read = reloaded.Sheets[0];
        Assert.Equal("Row 1 of a table long enough to need continuation records", read[1, 1].GetText());
        Assert.Equal("Row 2000 of a table long enough to need continuation records", read[2000, 1].GetText());
    }

    /// <summary>
    /// An independent reader has to make sense of the file too. Agreement between this
    /// library's own writer and reader proves only that they share assumptions.
    /// </summary>
    [Fact]
    public async Task AnIndependentReaderAgreesWithWhatWasWritten()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Shared");
        sheet["A1"] = CellValue.Text("Label");
        sheet["B1"] = CellValue.Number(42);
        sheet["A2"] = CellValue.Text("Другой язык");
        sheet["B2"] = CellValue.Number(3.5);
        sheet["A3"] = CellValue.Boolean(false);

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(buffer);
        Assert.Equal("Shared", reader.Name);

        Assert.True(reader.Read());
        Assert.Equal("Label", reader.GetValue(0));
        Assert.Equal(42.0, Convert.ToDouble(reader.GetValue(1), System.Globalization.CultureInfo.InvariantCulture));

        Assert.True(reader.Read());
        Assert.Equal("Другой язык", reader.GetValue(0));
        Assert.Equal(3.5, Convert.ToDouble(reader.GetValue(1), System.Globalization.CultureInfo.InvariantCulture));

        Assert.True(reader.Read());
        Assert.Equal(false, reader.GetValue(0));
    }

    /// <summary>
    /// A second independent reader, and a stricter one: NPOI parses every record it meets
    /// rather than skipping what it does not model, so it rejects a file whose record shapes
    /// are wrong even where ExcelDataReader would sail past.
    /// </summary>
    [Fact]
    public async Task AStricterReaderAgreesWithWhatWasWritten()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Sheet1");
        sheet["A1"] = CellValue.Text("Label");
        sheet["B1"] = CellValue.Number(42);
        sheet["A2"] = CellValue.DateTime(new DateTime(2024, 3, 17, 0, 0, 0, DateTimeKind.Unspecified));
        sheet.Cell("B2").SetFormula("B1*2");
        sheet.Values.Set(2, 2, CellValue.Number(84));
        sheet.Merge(RangeRef.Parse("C1:D1"));
        sheet.Validations.Add(DataValidation.WholeBetween("E1:E20", 1, 10));
        original.Names.Add("Total", "Sheet1!$B$1");

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        var oracle = new NPOI.HSSF.UserModel.HSSFWorkbook(buffer);
        NPOI.SS.UserModel.ISheet read = oracle.GetSheetAt(0);

        Assert.Equal("Sheet1", read.SheetName);
        Assert.Equal("Label", read.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(42, read.GetRow(0).GetCell(1).NumericCellValue);
        Assert.Equal(new DateTime(2024, 3, 17), read.GetRow(1).GetCell(0).DateCellValue);
        Assert.Equal("B1*2", read.GetRow(1).GetCell(1).CellFormula);
        Assert.Equal("C1:D1", Assert.Single(read.MergedRegions).FormatAsString());
        Assert.Single(read.GetDataValidations());
        Assert.Equal("Sheet1!$B$1", oracle.GetName("Total").RefersToFormula);
    }

    /// <summary>
    /// The sign of a unary term, checked against a reader that is not us. Writing and reading
    /// with the same wrong token numbers agrees with itself perfectly, so a round trip through
    /// this library cannot see the difference — and a swap there silently negates numbers both
    /// coming in and going out. Only an independent reader catches it.
    /// </summary>
    [Fact]
    public async Task AUnaryTermKeepsItsSignForAnotherReader()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Signs");
        sheet["A1"] = CellValue.Number(7);
        sheet.Cell("B1").SetFormula("-A1");
        sheet.Cell("B2").SetFormula("+A1");
        sheet.Cell("B3").SetFormula("-A1+A1");

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        var oracle = new NPOI.HSSF.UserModel.HSSFWorkbook(buffer);
        NPOI.SS.UserModel.ISheet read = oracle.GetSheetAt(0);

        Assert.Equal("-A1", read.GetRow(0).GetCell(1).CellFormula);
        Assert.Equal("+A1", read.GetRow(1).GetCell(1).CellFormula);
        Assert.Equal("-A1+A1", read.GetRow(2).GetCell(1).CellFormula);
    }

    /// <summary>
    /// A hyperlink and its tooltip survive being written and read back. The tooltip is the part
    /// worth checking against another reader: it lives in a record of its own after the link,
    /// and putting it inside the link structure instead would round-trip through this library
    /// perfectly while showing nothing in Excel.
    /// </summary>
    [Fact]
    public async Task AHyperlinkKeepsItsTooltipForAnotherReader()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Links");
        sheet["A1"] = CellValue.Text("Company");
        sheet.Hyperlinks.Add(RangeRef.Parse("A1"), "https://example.com/", "Opens the company site");
        sheet["A2"] = CellValue.Text("Elsewhere");
        sheet.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("A2"), "Links!B7"));

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        byte[] bytes = buffer.ToArray();

        // NPOI closes the stream it is handed, so each reader gets its own over the same bytes.
        using var forOracle = new MemoryStream(bytes, writable: false);
        var oracle = new NPOI.HSSF.UserModel.HSSFWorkbook(forOracle);
        NPOI.SS.UserModel.IHyperlink link = oracle.GetSheetAt(0).GetRow(0).GetCell(0).Hyperlink;
        Assert.Equal("https://example.com/", link.Address);

        using var forUs = new MemoryStream(bytes, writable: false);
        Workbook reread = await XlsWorkbook.LoadAsync(forUs, options: null, TestContext.Current.CancellationToken);
        Worksheet readSheet = reread.Sheets[0];

        Hyperlink external = Assert.Single(readSheet.Hyperlinks, static h => h.Target is not null);
        Assert.Equal("https://example.com/", external.Target);
        Assert.Equal("Opens the company site", external.Tooltip);

        Hyperlink internalLink = Assert.Single(readSheet.Hyperlinks, static h => h.Location is not null);
        Assert.Equal("Links!B7", internalLink.Location);
    }

    /// <summary>Anything the format cannot carry has to be reported, not dropped in silence.</summary>
    [Fact]
    public async Task ContentBeyondTheOldGridIsReported()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Big");
        sheet["A1"] = CellValue.Number(1);
        sheet[70_000, 1] = CellValue.Number(2);

        var warnings = new List<WorkbookWarning>();
        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(
            original, buffer, new XlsSaveOptions { OnWarning = warnings.Add }, TestContext.Current.CancellationToken);

        Assert.Contains(warnings, static warning => warning.Code == "GridTooLarge");
    }

    private static async Task<Workbook> RoundTripAsync(Workbook workbook)
    {
        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        return await XlsWorkbook.LoadAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }
}
