using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The print settings: how they come out of a real workbook, and how they survive being
/// written back. They are spread over a dozen records, so a round trip is the honest test.
/// </summary>
public sealed class PageSetupTests
{
    [Fact]
    public async Task PrintSettingsSurviveAWriteAndRead()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Report");
        sheet["A1"] = CellValue.Text("Title");

        PageSetup setup = sheet.PageSetup;
        setup.Orientation = PageOrientation.Landscape;
        setup.PaperSize = 9;
        setup.Scale = 85;
        setup.CenterHorizontally = true;
        setup.PrintGridLines = true;
        setup.PrintHeadings = true;
        setup.BlackAndWhite = true;
        setup.Order = PageOrder.OverThenDown;
        setup.FirstPageNumber = 3;
        setup.Margins = new PageMargins(0.5, 0.5, 1.25, 1.25, 0.4, 0.45);
        setup.Header = new HeaderFooterText { Default = "&CQuarterly report" };
        setup.Footer = new HeaderFooterText { Default = "&RPage &P of &N" };
        setup.AddRowBreak(20);
        setup.AddColumnBreak(5);

        Workbook reloaded = await RoundTripAsync(original);
        PageSetup read = reloaded.Sheets[0].PageSetup;

        Assert.Equal(PageOrientation.Landscape, read.Orientation);
        Assert.Equal(9, read.PaperSize);
        Assert.Equal(85, read.Scale);
        Assert.True(read.CenterHorizontally);
        Assert.False(read.CenterVertically);
        Assert.True(read.PrintGridLines);
        Assert.True(read.PrintHeadings);
        Assert.True(read.BlackAndWhite);
        Assert.Equal(PageOrder.OverThenDown, read.Order);
        Assert.Equal(3, read.FirstPageNumber);
        Assert.Equal(new PageMargins(0.5, 0.5, 1.25, 1.25, 0.4, 0.45), read.Margins);
        Assert.Equal("&CQuarterly report", read.Header?.Default);
        Assert.Equal("&RPage &P of &N", read.Footer?.Default);
        Assert.Equal([20], read.RowBreaks);
        Assert.Equal([5], read.ColumnBreaks);
    }

    /// <summary>
    /// Fitting to a page count lives in two records that only make sense together: the counts
    /// in one, the flag that turns them on in another, two blocks earlier.
    /// </summary>
    [Fact]
    public async Task FittingToPagesSurvives()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Wide");
        sheet["A1"] = CellValue.Number(1);
        sheet.PageSetup.FitToWidth = 1;
        sheet.PageSetup.FitToHeight = 2;

        PageSetup read = (await RoundTripAsync(original)).Sheets[0].PageSetup;

        Assert.True(read.FitToPage);
        Assert.Equal(1, read.FitToWidth);
        Assert.Equal(2, read.FitToHeight);
    }

    [Fact]
    public async Task DifferentFirstAndEvenHeadersSurvive()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Booklet");
        sheet["A1"] = CellValue.Number(1);
        sheet.PageSetup.DifferentOddEven = true;
        sheet.PageSetup.DifferentFirstPage = true;
        sheet.PageSetup.Header = new HeaderFooterText
        {
            Default = "&Codd",
            Even = "&Ceven",
            First = "&Cfirst",
        };

        PageSetup read = (await RoundTripAsync(original)).Sheets[0].PageSetup;

        Assert.True(read.DifferentOddEven);
        Assert.True(read.DifferentFirstPage);
        Assert.Equal("&Codd", read.Header?.Default);
        Assert.Equal("&Ceven", read.Header?.Even);
        Assert.Equal("&Cfirst", read.Header?.First);
    }

    /// <summary>
    /// Real workbooks, to check the reader against records this library did not write. Excel
    /// leaves out the margin records when they are its defaults, so the corpus is scanned for
    /// the files that do state them rather than demanding them of any one workbook.
    /// </summary>
    [Fact]
    public async Task PrintSettingsOfRealWorkbooksAreRead()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int withMargins = 0;
        int withHeaders = 0;
        int examined = 0;

        foreach (string path in XlsCorpus.Files.Take(120))
        {
            Workbook workbook;
            try
            {
                workbook = await XlsWorkbook.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            foreach (Worksheet sheet in workbook.Sheets)
            {
                examined++;
                PageSetup setup = sheet.PageSetup;
                Assert.InRange(setup.Scale, 10, 400);
                if (setup.Margins is { } margins)
                {
                    withMargins++;
                    Assert.InRange(margins.Left, 0, 49);
                    Assert.InRange(margins.Header, 0, 49);
                }

                if (setup.Header?.Default is { Length: > 0 } or null && setup.Footer?.Default is { Length: > 0 })
                    withHeaders++;
            }
        }

        Assert.True(examined > 50, $"Only {examined} sheets were readable.");
        Assert.True(withMargins > 5, $"Only {withMargins} sheets stated their margins; the records may be unread.");
        Assert.True(withHeaders > 0, "No sheet had a footer, which a hundred real workbooks should not manage.");
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
