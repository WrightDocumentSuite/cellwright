using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// The everyday sheet features added on top of cells: row and column visibility, outline
/// grouping, tab appearance, hyperlinks, auto filter and input rules. Every case goes through
/// a real package so that both halves of the round trip are exercised.
/// </summary>
public sealed class WorksheetFeatureTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task EverySheetFeature_ReadsBackThroughClosedXml()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Everything");
        sheet["A1"] = "Region";
        sheet["B1"] = "Amount";
        sheet["A2"] = "North";
        sheet["B2"] = 100;
        sheet["F1"] = "Code";
        sheet["G1"] = "Label";
        sheet["F2"] = 1;
        sheet["G2"] = "one";

        sheet.Row(3).Hide();
        sheet.Column("D").SetWidth(30);
        sheet.TabColor = ExcelColor.FromRgb(0x22, 0x78, 0xD4);
        sheet.AutoFilter("A1:B2");
        sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com");
        sheet.Comments.Add("B2", "Verified", "QA");
        sheet.Validations.Add(DataValidation.List("C2:C50", "\"Yes,No\""));
        sheet.Tables.Add("Codes", "F1:G2");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        IXLWorksheet loaded = interop.Worksheet("Everything");
        Assert.True(loaded.Row(3).IsHidden);
        Assert.True(loaded.AutoFilter.IsEnabled);
        Assert.Equal("https://example.com/", loaded.Cell("A2").GetHyperlink().ExternalAddress.ToString());
        Assert.Equal("Verified", loaded.Cell("B2").GetComment().Text);
        Assert.Equal(XLAllowedValues.List, loaded.Cell("C2").GetDataValidation().AllowedValues);
        Assert.Equal("Codes", Assert.Single(loaded.Tables).Name);
    }

    [Fact]
    public async Task RowsAndColumns_VisibilityHeightsAndOutline_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        sheet["A1"] = "top";
        sheet["A20"] = "bottom";

        sheet.Row(1).SetHeight(28);
        sheet.Row(3).Hide();
        sheet.Rows(4, 6).Group();
        sheet.Row(7).Collapsed = true;
        sheet.Column("B").SetWidth(18).Hide();
        sheet.Columns(4, 6).Group(collapse: true);

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(28, loaded.Row(1).Height);
        Assert.True(loaded.Row(3).Hidden);
        Assert.Equal(1, loaded.Row(5).OutlineLevel);
        Assert.True(loaded.Row(7).Collapsed);
        Assert.Equal(18, loaded.Column(2).Width);
        Assert.True(loaded.Column(2).Hidden);
        Assert.Equal(1, loaded.Column(5).OutlineLevel);
        Assert.True(loaded.Column(5).Hidden);
        Assert.True(loaded.Column(7).Collapsed);
    }

    [Fact]
    public async Task RowConfiguredBeyondUsedRange_Survives()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sparse");
        sheet["A1"] = 1;
        sheet.Row(500).Hide();

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Sparse"];
        Assert.True(loaded.Row(500).Hidden);
    }

    [Fact]
    public async Task AdjacentIdenticalColumns_CollapseIntoOneSpan()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Wide");
        sheet["A1"] = 1;
        sheet.Columns(2, 40).SetWidth(9.5);

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Wide"];
        Assert.Equal(9.5, loaded.Column(2).Width);
        Assert.Equal(9.5, loaded.Column(40).Width);
        Assert.Null(loaded.Column(41).Width);
    }

    [Fact]
    public async Task SheetAppearance_TabColorGridLinesAndZoom_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Look");
        sheet["A1"] = 1;
        sheet.TabColor = ExcelColor.FromArgb(0xFF2278D4);
        sheet.ShowGridLines = false;
        sheet.ZoomScale = 145;

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Look"];
        Assert.Equal(0xFF2278D4u, loaded.TabColor.Argb);
        Assert.False(loaded.ShowGridLines);
        Assert.Equal(145, loaded.ZoomScale);
    }

    [Fact]
    public async Task Hyperlinks_ExternalAndInternal_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Links");
        Worksheet other = book.Sheets.Add("Target");
        other["A1"] = "here";

        sheet["A1"] = "Cellwright";
        sheet["A2"] = "Jump";
        sheet.Hyperlinks.Add(RangeRef.Parse("A1"), "https://example.com/docs", "Open the docs");
        sheet.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("A2"), "Target!A1"));

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Links"];
        Assert.Equal(2, loaded.Hyperlinks.Count);

        Hyperlink external = loaded.Hyperlinks.Find(CellRef.Parse("A1"))!;
        Assert.Equal("https://example.com/docs", external.Target);
        Assert.Equal("Open the docs", external.Tooltip);
        Assert.True(external.IsExternal);

        Hyperlink internalLink = loaded.Hyperlinks.Find(CellRef.Parse("A2"))!;
        Assert.False(internalLink.IsExternal);
        Assert.Equal("Target!A1", internalLink.Location);
    }

    [Fact]
    public async Task AutoFilterAndBuiltInNames_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Table");
        sheet["A1"] = "Name";
        sheet["B1"] = "Amount";
        sheet["A2"] = "x";
        sheet["B2"] = 10;
        sheet.AutoFilter("A1:B2");
        book.Names.Add("_xlnm.Print_Area", "Table!$A$1:$B$2", sheet);

        Workbook reloaded = await SaveAndReloadAsync(book);
        Assert.Equal("A1:B2", reloaded.Sheets["Table"].AutoFilterRange!.Value.ToString());

        DefinedName printArea = Assert.Single(reloaded.Names);
        Assert.True(printArea.IsBuiltIn);
        Assert.Equal("Table!$A$1:$B$2", printArea.RefersTo);
        Assert.Same(reloaded.Sheets["Table"], printArea.Scope);
    }

    [Fact]
    public async Task Comments_TextAndAuthor_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Notes");
        sheet["A1"] = "value";
        sheet.Comments.Add("A1", "Check this figure", "Reviewer");
        sheet.Comments.Add("C3", "Второй комментарий");

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Notes"];
        Assert.Equal(2, loaded.Comments.Count);

        CellComment first = loaded.Comments[CellRef.Parse("A1")]!;
        Assert.Equal("Check this figure", first.Text);
        Assert.Equal("Reviewer", first.Author);
        Assert.Equal("Второй комментарий", loaded.Comments[CellRef.Parse("C3")]!.Text);
    }

    [Fact]
    public async Task Comments_RemovingTheLastOne_DropsThePart()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Notes");
        sheet["A1"] = "value";
        sheet.Comments.Add("A1", "temporary");

        Workbook reloaded = await SaveAndReloadAsync(book);
        Worksheet loaded = reloaded.Sheets["Notes"];
        Assert.Single(loaded.Comments);

        Assert.True(loaded.Comments.Remove(CellRef.Parse("A1")));
        Assert.Empty((await SaveAndReloadAsync(reloaded)).Sheets["Notes"].Comments);
    }

    [Fact]
    public async Task Tables_ColumnsAndStyle_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sales");
        sheet["A1"] = "Region";
        sheet["B1"] = "Amount";
        sheet["A2"] = "North";
        sheet["B2"] = 100;
        sheet["A3"] = "South";
        sheet["B3"] = 250;

        sheet.Tables.Add("SalesTable", "A1:B3");

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Sales"];
        ExcelTable table = Assert.Single(loaded.Tables);
        Assert.Equal("SalesTable", table.Name);
        Assert.Equal("A1:B3", table.Range.ToString());
        Assert.Equal(["Region", "Amount"], table.Columns);
        Assert.True(table.HasHeaderRow);
    }

    [Fact]
    public async Task Tables_BlankHeaderCells_GetGeneratedNames()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Gaps");
        sheet["A1"] = "Region";
        sheet["A2"] = "North";

        sheet.Tables.Add("Gappy", "A1:C2");

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Gaps"];
        Assert.Equal(["Region", "Column2", "Column3"], Assert.Single(loaded.Tables).Columns);
    }

    [Fact]
    public async Task DataValidations_ListAndBounds_RoundTrip()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Input");
        sheet["A1"] = "pick";

        sheet.Validations.Add(DataValidation.List("A2:A100", "\"Yes,No\""));
        sheet.Validations.Add(new DataValidation
        {
            Ranges = [RangeRef.Parse("B2:B100"), RangeRef.Parse("D2:D100")],
            Type = DataValidationType.Whole,
            Operator = DataValidationOperator.Between,
            Formula1 = "1",
            Formula2 = "99",
            ErrorTitle = "Out of range",
            Error = "Use a number from 1 to 99.",
            ErrorStyle = DataValidationErrorStyle.Warning,
        });

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Input"];
        Assert.Equal(2, loaded.Validations.Count);

        DataValidation list = loaded.Validations[0];
        Assert.Equal(DataValidationType.List, list.Type);
        Assert.Equal("\"Yes,No\"", list.Formula1);
        Assert.Equal("A2:A100", Assert.Single(list.Ranges).ToString());

        DataValidation whole = loaded.Validations[1];
        Assert.Equal(DataValidationType.Whole, whole.Type);
        Assert.Equal("99", whole.Formula2);
        Assert.Equal(2, whole.Ranges.Count);
        Assert.Equal(DataValidationErrorStyle.Warning, whole.ErrorStyle);
        Assert.Equal("Out of range", whole.ErrorTitle);
    }
}
