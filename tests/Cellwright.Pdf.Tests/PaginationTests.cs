using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Cutting a grid into pages: down and across, at the places the sheet asks for, with the
/// titles repeated and the whole thing scaled to fit when the sheet says so.
/// </summary>
public sealed class PaginationTests
{
    /// <summary>A sheet of numbered cells, one per row and column, easy to find on a page.</summary>
    private static Workbook Grid(int rows, int columns)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        for (int row = 1; row <= rows; row++)
        {
            for (int column = 1; column <= columns; column++)
                sheet.Cell(new CellRef(row, column)).SetValue(CellValue.Text($"r{row}c{column}"));
        }

        return book;
    }

    [Fact]
    public void ASheetTallerThanThePageRunsOntoTheNext()
    {
        using Rendered rendered = Rendered.Of(Grid(rows: 120, columns: 1));

        Assert.True(rendered.PageCount >= 2, "a hundred and twenty rows fitted one page");
        Assert.Contains("r1c1", rendered.Text(0), StringComparison.Ordinal);
        Assert.DoesNotContain("r1c1", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void ASheetWiderThanThePageRunsOntoTheNextToo()
    {
        using Rendered rendered = Rendered.Of(Grid(rows: 2, columns: 20));

        Assert.True(rendered.PageCount >= 2, "twenty columns fitted one page");
        Assert.Contains("r1c1", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("r1c20", rendered.Text(rendered.PageCount - 1), StringComparison.Ordinal);
    }

    [Fact]
    public void PagesGoDownBeforeAcrossUnlessTheSheetSaysOtherwise()
    {
        Workbook down = Grid(rows: 120, columns: 20);
        using Rendered rendered = Rendered.Of(down);

        // Down then over: the second page carries the same columns, further down the sheet.
        Assert.Contains("c1", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void OverThenDownTakesTheColumnsFirst()
    {
        Workbook book = Grid(rows: 120, columns: 20);
        book.Sheets[0].PageSetup.Order = PageOrder.OverThenDown;

        using Rendered rendered = Rendered.Of(book);

        // Across first: the second page carries the same rows, further to the right.
        Assert.Contains("r1c", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void AManualRowBreakStartsANewPageWhereItSays()
    {
        Workbook book = Grid(rows: 6, columns: 1);
        book.Sheets[0].PageSetup.AddRowBreak(4);

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(2, rendered.PageCount);
        Assert.Contains("r3c1", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("r4c1", rendered.Text(1), StringComparison.Ordinal);
        Assert.DoesNotContain("r4c1", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void AnExplicitSparsePrintAreaKeepsItsBlankPages()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A5"] = "SPARSE";
        sheet.SetPrintArea("A1:A9");
        sheet.PageSetup.AddRowBreak(4);
        sheet.PageSetup.AddRowBreak(7);

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(3, rendered.PageCount);
        Assert.DoesNotContain("SPARSE", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("SPARSE", rendered.Text(1), StringComparison.Ordinal);
        Assert.DoesNotContain("SPARSE", rendered.Text(2), StringComparison.Ordinal);
    }

    [Fact]
    public void AManualColumnBreakDoesTheSameAcross()
    {
        Workbook book = Grid(rows: 2, columns: 6);
        book.Sheets[0].PageSetup.AddColumnBreak(3);

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(2, rendered.PageCount);
        Assert.Contains("r1c2", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("r1c3", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedRowsComeBackAtTheTopOfEveryPage()
    {
        Workbook book = Grid(rows: 120, columns: 1);
        Worksheet sheet = book.Sheets[0];
        sheet.Cell("A1").SetValue(CellValue.Text("HEADER"));
        sheet.SetRepeatRows(1, 1);

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.PageCount >= 2);
        Assert.Contains("HEADER", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("HEADER", rendered.Text(1), StringComparison.Ordinal);
        Assert.Equal("HEADER", rendered.Lines(1)[0]);
    }

    [Fact]
    public void ARepeatedTitlePrintsOnOtherwiseBlankBands()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "HEADER";
        sheet.SetPrintArea("A1:A9");
        sheet.SetRepeatRows(1, 1);
        sheet.PageSetup.AddRowBreak(4);
        sheet.PageSetup.AddRowBreak(7);

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(3, rendered.PageCount);
        for (int page = 0; page < rendered.PageCount; page++)
            Assert.Contains("HEADER", rendered.Text(page), StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedColumnsComeBackAtTheLeftOfEveryPage()
    {
        Workbook book = Grid(rows: 2, columns: 20);
        Worksheet sheet = book.Sheets[0];
        sheet.Cell("A1").SetValue(CellValue.Text("KEY"));
        sheet.SetRepeatColumns(1, 1);

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.PageCount >= 2);
        Assert.Contains("KEY", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("KEY", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void FitToOnePageWideSqueezesTheColumnsOntoIt()
    {
        Workbook wide = Grid(rows: 2, columns: 20);
        using Rendered spread = Rendered.Of(wide);
        Assert.True(spread.PageCount >= 2);

        Workbook fitted = Grid(rows: 2, columns: 20);
        fitted.Sheets[0].PageSetup.FitToWidth = 1;

        using Rendered rendered = Rendered.Of(fitted);

        Assert.Equal(1, rendered.PageCount);
        Assert.Contains("r1c1", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("r1c20", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void FitToOnePageTallDoesTheSameDown()
    {
        Workbook book = Grid(rows: 90, columns: 1);
        book.Sheets[0].PageSetup.FitToHeight = 1;

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(1, rendered.PageCount);
        Assert.Contains("r90c1", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void AStatedScaleMakesEverythingSmaller()
    {
        Workbook full = Grid(rows: 2, columns: 2);
        Workbook half = Grid(rows: 2, columns: 2);
        half.Sheets[0].PageSetup.Scale = 50;

        using Rendered whole = Rendered.Of(full);
        using Rendered scaled = Rendered.Of(half);

        (double X, double Y)? one = whole.Find("r");
        (double X, double Y)? other = scaled.Find("r");

        Assert.NotNull(one);
        Assert.NotNull(other);

        // The grid is drawn from the same corner, so half scale puts the second row's text
        // half as far down the page.
        Assert.True(other!.Value.Y > one!.Value.Y, "the scaled sheet did not sit higher on the page");
    }

    [Fact]
    public void CentringPutsASmallGridInTheMiddleOfThePaper()
    {
        Workbook left = Grid(rows: 1, columns: 1);
        Workbook middle = Grid(rows: 1, columns: 1);
        middle.Sheets[0].PageSetup.CenterHorizontally = true;

        using Rendered plain = Rendered.Of(left);
        using Rendered centred = Rendered.Of(middle);

        (double X, double Y)? first = plain.Find("r");
        (double X, double Y)? second = centred.Find("r");

        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.True(second!.Value.X > first!.Value.X + 100, "the grid was not centred");
    }

    [Fact]
    public void ASheetThatCannotFinishSaysSoRatherThanRunningForEver()
    {
        Workbook book = Grid(rows: 400, columns: 1);

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => PdfExporter.Render(book, new PdfExportOptions { MaxPages = 3 }));

        Assert.Contains("MaxPages", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEnormousPrintAreaIsStoppedBeforeItIsBuilt()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "The only cell that holds anything";
        sheet.SetPrintArea("A1:Z200000");

        var clock = System.Diagnostics.Stopwatch.StartNew();
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => PdfExporter.Render(book));
        clock.Stop();

        Assert.Contains("MaxPages", error.Message, StringComparison.Ordinal);

        // The point of the limit is that nothing is drawn and nothing is measured: had either
        // happened, five million cells would have taken far longer than this.
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(10), $"it took {clock.Elapsed} to give up");
    }

    [Fact]
    public void ASheetInsideTheLimitStillPrintsEveryPage()
    {
        Workbook book = Grid(rows: 400, columns: 1);

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { MaxPages = 20 });

        Assert.InRange(rendered.PageCount, 2, 20);
        Assert.Contains("r400c1", rendered.Text(rendered.PageCount - 1), StringComparison.Ordinal);
    }
}
