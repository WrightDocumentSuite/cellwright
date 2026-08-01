using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

public sealed class WorkbookTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        Assert.Empty(new OpenXmlValidator().Validate(document, TestContext.Current.CancellationToken));

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task CreateEditSaveLoad_FullRoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Отчёт");

        sheet["A1"] = "Товар";
        sheet["B1"] = "Цена";
        sheet[2, 1] = "Яблоко";
        sheet[2, 2] = 2.99;
        sheet["A3"] = "Апельсин";
        sheet["B3"] = 1.99;
        sheet["A4"] = new DateTime(2026, 7, 24, 12, 0, 0);
        sheet["B4"] = true;
        sheet.Cell("B5").SetFormula("SUM(B2:B3)");
        sheet.Range("A1:B1").Merge();
        sheet.SetColumnWidth(1, 22.5);
        sheet.SetRowHeight(1, 30);
        sheet.FrozenRows = 1;

        Worksheet hidden = book.Sheets.Add("Скрытый");
        hidden["A1"] = 1;
        hidden.IsHidden = true;

        Workbook reloaded = await SaveAndReloadAsync(book);

        Worksheet loaded = reloaded.Sheets["Отчёт"];
        Assert.Equal("Товар", loaded["A1"].GetText());
        Assert.Equal(2.99, loaded["B2"].GetNumber());
        Assert.Equal(new DateTime(2026, 7, 24, 12, 0, 0), loaded["A4"].GetDateTime());
        Assert.True(loaded["B4"].GetBoolean());
        Assert.Equal("SUM(B2:B3)", loaded.Cell("B5").Formula);
        Assert.Equal("A1:B1", Assert.Single(loaded.MergedRanges).ToString());
        Assert.Equal(22.5, loaded.Column(1).Width);
        Assert.Equal(30, loaded.Row(1).Height);
        Assert.Equal(1, loaded.FrozenRows);
        Assert.True(reloaded.Sheets["Скрытый"].IsHidden);
        Assert.False(reloaded.LoadDiagnostics!.HasWarnings);
    }

    [Fact]
    public async Task Styles_SurviveRoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        var headerStyle = new CellStyle
        {
            Font = new FontStyle { Bold = true, Size = 13 },
            Fill = FillStyle.Solid(ExcelColor.FromArgb(0xFFEEDD00)),
            Alignment = new Alignment { Horizontal = HorizontalAlignment.Center },
            NumberFormat = "#,##0.000",
        };
        sheet.Cell("A1").SetValue(1234.5678).SetStyle(headerStyle);

        Workbook reloaded = await SaveAndReloadAsync(book);
        CellStyle loaded = reloaded.Sheets["S"].Cell("A1").StyleValue;

        Assert.True(loaded.Font.Bold);
        Assert.Equal(13, loaded.Font.Size);
        Assert.True(loaded.Fill.IsSolid);
        Assert.Equal(0xFFEEDD00, loaded.Fill.Color.Argb);
        Assert.Equal(HorizontalAlignment.Center, loaded.Alignment.Horizontal);
        Assert.Equal("#,##0.000", loaded.NumberFormat);
    }

    [Fact]
    public async Task FormulaWorkbook_TriggersFullRecalcFlag()
    {
        var book = Workbook.Create();
        book.Sheets.Add("S").Cell("A1").SetFormula("=1+2");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
        string workbookXml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
        Assert.Contains("fullCalcOnLoad=\"1\"", workbookXml, StringComparison.Ordinal);
    }

    [Fact]
    public void CellFacade_ValueAndFormulaExcludeEachOther()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");

        Cell cell = sheet.Cell("A1").SetFormula("SUM(B1:B9)");
        Assert.Equal("SUM(B1:B9)", cell.Formula);
        Assert.True(cell.Value.IsBlank);

        cell.SetValue(42);
        Assert.Null(cell.Formula);
        Assert.Equal(42.0, cell.Value.GetNumber());
    }

    [Fact]
    public void Range_BulkOperations()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        StyleId style = book.RegisterStyle(new CellStyle { Font = new FontStyle { Italic = true } });

        sheet.Range("A1:C2").SetValue(7).SetStyle(style);

        Assert.Equal(7.0, sheet["C2"].GetNumber());
        Assert.Equal(style, sheet.Cell("B1").Style);
        Assert.Equal(new RangeRef(new CellRef(1, 1), new CellRef(2, 3)), sheet.UsedRange);

        sheet.Range("A1:C2").Clear();
        Assert.True(sheet["A1"].IsBlank);
    }

    [Fact]
    public async Task LoadsClosedXmlWorkbook_AndReadsCellsAndStyles()
    {
        using var stream = new MemoryStream();
        using (var external = new XLWorkbook())
        {
            IXLWorksheet ws = external.AddWorksheet("Data");
            ws.Cell(1, 1).Value = "bold";
            ws.Cell(1, 1).Style.Font.SetBold(true);
            ws.Cell(2, 2).Value = 99.5;
            ws.Cell(3, 1).FormulaA1 = "=B2*2";
            ws.Range("A5:B5").Merge();
            external.SaveAs(stream);
        }

        stream.Position = 0;
        Workbook book = await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Worksheet sheet = book.Sheets["Data"];

        Assert.Equal("bold", sheet["A1"].GetText());
        Assert.True(sheet.Cell("A1").StyleValue.Font.Bold);
        Assert.Equal(99.5, sheet["B2"].GetNumber());
        Assert.Equal("B2*2", sheet.Cell("A3").Formula);
        Assert.Equal("A5:B5", Assert.Single(sheet.MergedRanges).ToString());
    }

    [Fact]
    public async Task SavedWorkbook_OpensInClosedXml()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Interop");
        sheet["A1"] = "from Cellwright";
        sheet["B1"] = 3.5;
        sheet.Cell("C1").SetFormula("B1*2");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        IXLWorksheet ws = workbook.Worksheet("Interop");
        Assert.Equal("from Cellwright", ws.Cell(1, 1).GetString());
        Assert.Equal(3.5, ws.Cell(1, 2).GetDouble());
        Assert.Equal("B1*2", ws.Cell(1, 3).FormulaA1);
    }

    [Fact]
    public void SheetManagement_AddFindRemoveRename()
    {
        var book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        book.Sheets.Add("Second");

        Assert.Throws<ArgumentException>(() => book.Sheets.Add("first"));
        Assert.Throws<ArgumentException>(() => first.Name = "SECOND");

        first.Name = "Renamed";
        Assert.NotNull(book.Sheets.Find("renamed"));
        Assert.True(book.Sheets.Remove("Second"));
        Assert.Single(book.Sheets);
        Assert.Throws<KeyNotFoundException>(() => book.Sheets["missing"]);
    }
}
