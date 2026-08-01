using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Print settings: paper and scaling, margins, headers, page breaks, and the built-in names
/// behind the print area and the repeated titles.
/// </summary>
public sealed class PageSetupTests
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

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        sheet["A1"] = "Region";
        sheet["A2"] = "North";
        return book;
    }

    [Fact]
    public async Task PaperScalingAndPrintOptionsRoundTrip()
    {
        Workbook book = Sample();
        PageSetup setup = book.Sheets[0].PageSetup;
        setup.Orientation = PageOrientation.Landscape;
        setup.PaperSize = 9;
        setup.Scale = 80;
        setup.PrintGridLines = true;
        setup.PrintHeadings = true;
        setup.CenterHorizontally = true;
        setup.BlackAndWhite = true;
        setup.FirstPageNumber = 4;
        setup.Order = PageOrder.OverThenDown;

        PageSetup loaded = (await SaveAndReloadAsync(book)).Sheets[0].PageSetup;

        Assert.Equal(PageOrientation.Landscape, loaded.Orientation);
        Assert.Equal(9, loaded.PaperSize);
        Assert.Equal(80, loaded.Scale);
        Assert.True(loaded.PrintGridLines);
        Assert.True(loaded.PrintHeadings);
        Assert.True(loaded.CenterHorizontally);
        Assert.False(loaded.CenterVertically);
        Assert.True(loaded.BlackAndWhite);
        Assert.Equal(4, loaded.FirstPageNumber);
        Assert.Equal(PageOrder.OverThenDown, loaded.Order);
    }

    [Fact]
    public async Task FitToPageSurvivesThroughSheetProperties()
    {
        Workbook book = Sample();
        book.Sheets[0].PageSetup.FitToWidth = 1;
        book.Sheets[0].PageSetup.FitToHeight = null;

        PageSetup loaded = (await SaveAndReloadAsync(book)).Sheets[0].PageSetup;

        Assert.True(loaded.FitToPage);
        Assert.Equal(1, loaded.FitToWidth);
        Assert.Null(loaded.FitToHeight);
    }

    [Fact]
    public async Task MarginsAndHeadersRoundTrip()
    {
        Workbook book = Sample();
        PageSetup setup = book.Sheets[0].PageSetup;
        setup.Margins = PageMargins.Narrow;
        setup.DifferentFirstPage = true;
        setup.Header = new HeaderFooterText { Default = "&CQuarterly report", First = "&CCover" };
        setup.Footer = new HeaderFooterText { Default = "&RPage &P of &N" };

        PageSetup loaded = (await SaveAndReloadAsync(book)).Sheets[0].PageSetup;

        Assert.Equal(PageMargins.Narrow, loaded.Margins);
        Assert.True(loaded.DifferentFirstPage);
        Assert.Equal("&CQuarterly report", loaded.Header?.Default);
        Assert.Equal("&CCover", loaded.Header?.First);
        Assert.Equal("&RPage &P of &N", loaded.Footer?.Default);
    }

    [Fact]
    public async Task PageBreaksRoundTrip()
    {
        Workbook book = Sample();
        book.Sheets[0].PageSetup.AddRowBreak(25);
        book.Sheets[0].PageSetup.AddRowBreak(50);
        book.Sheets[0].PageSetup.AddColumnBreak(5);

        PageSetup loaded = (await SaveAndReloadAsync(book)).Sheets[0].PageSetup;

        Assert.Equal([25, 50], loaded.RowBreaks);
        Assert.Equal([5], loaded.ColumnBreaks);
    }

    [Fact]
    public async Task PrintAreaAndRepeatTitlesRoundTrip()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.SetPrintArea("A1:H40");
        sheet.SetRepeatRows(1, 2);
        sheet.SetRepeatColumns(1, 1);

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets[0];

        Assert.Equal(RangeRef.Parse("A1:H40"), loaded.PrintArea);
        Assert.Equal((1, 2), loaded.RepeatRows);
        Assert.Equal((1, 1), loaded.RepeatColumns);
    }

    [Fact]
    public async Task ClosedXmlAgreesOnTheSettings()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.PageSetup.Orientation = PageOrientation.Landscape;
        sheet.PageSetup.Margins = PageMargins.Wide;
        sheet.PageSetup.AddRowBreak(10);
        sheet.SetPrintArea("A1:D20");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        IXLWorksheet loaded = interop.Worksheet("Report");
        Assert.Equal(XLPageOrientation.Landscape, loaded.PageSetup.PageOrientation);
        Assert.Equal(1.0, loaded.PageSetup.Margins.Left, precision: 6);
        Assert.Equal(9, Assert.Single(loaded.PageSetup.RowBreaks));
        Assert.Equal("A1:D20", Assert.Single(loaded.PageSetup.PrintAreas).RangeAddress.ToStringRelative());
    }

    [Fact]
    public async Task UntouchedSettingsAreReplayedByteForByte()
    {
        Workbook book = Sample();
        book.Sheets[0].PageSetup.Orientation = PageOrientation.Portrait;
        Workbook reloaded = await SaveAndReloadAsync(book);

        // Nothing was touched after loading, so the second save must not invent settings.
        reloaded.Sheets[0]["A3"] = "South";
        Workbook again = await SaveAndReloadAsync(reloaded);

        Assert.Equal(PageOrientation.Portrait, again.Sheets[0].PageSetup.Orientation);
        Assert.False(again.Sheets[0].PageSetup.IsModified);
    }
}
