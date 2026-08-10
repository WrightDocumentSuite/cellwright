using Cellwright.Styles;
using Inkwright.Text;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// The grid on the page: what each cell shows, where in its cell it sits, what is painted behind
/// it and what is ruled around it.
/// </summary>
public sealed class GridTests
{
    private static Worksheet Sheet(Workbook book) => book.Sheets.Add("Sheet1");

    [Fact]
    public void ValuesAreShownTheWayTheirFormatSays()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = 1234.5;
        sheet.Cell("A1").SetStyle(CellStyle.Default with { NumberFormat = "#,##0.00" });
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("1,234.50", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void ADateIsShownAsADateRatherThanItsSerialNumber()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = new DateTime(2026, 3, 14, 0, 0, 0, DateTimeKind.Unspecified);
        sheet.Cell("A1").SetStyle(CellStyle.Default with { NumberFormat = "yyyy-mm-dd" });
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("2026-03-14", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void NumbersGoRightAndTextGoesLeftWhenNothingSaysOtherwise()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "ab";
        sheet["A2"] = 7d;
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);

        (double X, double Y)? text = rendered.Find("a");
        (double X, double Y)? number = rendered.Find("7");

        Assert.NotNull(text);
        Assert.NotNull(number);
        Assert.True(number!.Value.X > text!.Value.X + 40, "the number was not pushed to the right of its cell");
    }

    [Fact]
    public void AStatedAlignmentBeatsTheOneTheValueWouldGet()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = 7d;
        sheet["A2"] = 7d;
        sheet.Cell("A2").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { Horizontal = HorizontalAlignment.Left },
        });
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> sevens = [.. rendered.Letters().Where(letter => letter.Text == "7")];

        Assert.Equal(2, sevens.Count);
        Assert.True(
            sevens.Min(letter => letter.Origin.X) < sevens.Max(letter => letter.Origin.X) - 40,
            "the left-aligned number sat where the right-aligned one did");
    }

    [Fact]
    public void AFillIsPaintedBehindItsCell()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Filled";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Fill = FillStyle.Solid(ExcelColor.FromRgb(255, 0, 0)),
        });

        using Rendered rendered = Rendered.Of(book);
        string content = rendered.Content();

        Assert.Contains("1 0 0 rg", content, StringComparison.Ordinal);
        Assert.Contains(" re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ABorderIsStrokedRoundItsCell()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Boxed";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Border = BorderStyle.All(new BorderEdge(BorderLineStyle.Thin, ExcelColor.FromRgb(0, 0, 255))),
        });

        using Rendered rendered = Rendered.Of(book);
        string content = rendered.Content();

        Assert.Contains("0 0 1 RG", content, StringComparison.Ordinal);
        Assert.Contains("\nS\n", content, StringComparison.Ordinal);
    }

    [Fact]
    public void AStyleOnABlankCellStillPaintsItsFillAndBorder()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet.SetPrintArea("A1:A3");
        sheet.Cell("A2").SetStyle(CellStyle.Default with
        {
            Fill = FillStyle.Solid(ExcelColor.FromRgb(255, 0, 0)),
            Border = BorderStyle.All(new BorderEdge(BorderLineStyle.Thin, ExcelColor.FromRgb(0, 0, 255))),
        });

        using Rendered rendered = Rendered.Of(book);
        string content = rendered.Content();

        Assert.Contains("1 0 0 rg", content, StringComparison.Ordinal);
        Assert.Contains("0 0 1 RG", content, StringComparison.Ordinal);
    }

    [Fact]
    public void GridLinesStillPrintAcrossAnOtherwiseBlankArea()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet.SetPrintArea("A1:B2");
        sheet.PageSetup.PrintGridLines = true;

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("\nS\n", rendered.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheGridIsPrintedOnlyWhenTheSheetAsksForIt()
    {
        Workbook plain = Workbook.Create();
        Sheet(plain)["A1"] = "x";

        Workbook gridded = Workbook.Create();
        Worksheet sheet = Sheet(gridded);
        sheet["A1"] = "x";
        sheet.PageSetup.PrintGridLines = true;

        using Rendered without = Rendered.Of(plain);
        using Rendered with = Rendered.Of(gridded);

        Assert.DoesNotContain("\nS\n", without.Content(), StringComparison.Ordinal);
        Assert.Contains("\nS\n", with.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void AMergedCellIsDrawnOnceAcrossItsWholeRange()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Merged across";
        sheet["A2"] = "Below";
        sheet.Merge(RangeRef.Parse("A1:C1"));
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { Horizontal = HorizontalAlignment.Center },
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Merged across", rendered.Text(0), StringComparison.Ordinal);

        // Centred across three columns, the text starts well right of a single column's edge.
        (double X, double Y)? merged = rendered.Find("M");
        Assert.NotNull(merged);
        Assert.True(merged!.Value.X > 80, $"the merge was not spread across its columns: {merged.Value.X:0.0}");
    }

    [Fact]
    public void AMergeTheSizeOfTheSheetCostsNoMoreThanThePageItIsSeenOn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "A heading over a merge as tall as the sheet";
        sheet.Merge(RangeRef.Parse("A1:C1000000"));
        sheet.SetPrintArea("A1:C3");

        var clock = System.Diagnostics.Stopwatch.StartNew();
        using Rendered rendered = Rendered.Of(book);
        clock.Stop();

        Assert.Contains("A heading", rendered.Text(0), StringComparison.Ordinal);

        // Three million cells were covered by that merge and three of them are on the paper.
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(20), $"it took {clock.Elapsed} to draw one page");
    }

    [Fact]
    public void ANumberTooWideForItsColumnPrintsAsHashes()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = 123456789012345d;
        sheet.Cell("A1").SetStyle(CellStyle.Default with { NumberFormat = "#,##0.00" });
        sheet.Column("A").SetWidth(4);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("#", rendered.Text(0), StringComparison.Ordinal);
        Assert.DoesNotContain("123456", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void TextTooWideForItsColumnIsNotTurnedIntoHashes()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "A sentence far longer than its column";
        sheet.Column("A").SetWidth(4);

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain("#", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("A sentence", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void AHiddenRowIsNotPrintedAndTakesNoRoom()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "First";
        sheet["A2"] = "Hidden";
        sheet["A3"] = "Third";
        sheet.Row(2).Hidden = true;

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<string> lines = rendered.Lines();

        Assert.DoesNotContain("Hidden", rendered.Text(0), StringComparison.Ordinal);
        Assert.Equal(2, lines.Count);
    }

    [Fact]
    public void OnlyThePrintAreaIsPrintedWhenTheSheetNamesOne()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Inside";
        sheet["A5"] = "Outside";
        sheet.SetPrintArea("A1:B3");

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Inside", rendered.Text(0), StringComparison.Ordinal);
        Assert.DoesNotContain("Outside", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void RowsStackDownThePageInOrder()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "One";
        sheet["A2"] = "Two";
        sheet["A3"] = "Three";

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<string> lines = rendered.Lines();

        Assert.Equal(["One", "Two", "Three"], lines);
    }
}
