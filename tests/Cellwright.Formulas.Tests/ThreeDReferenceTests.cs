using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// 3-D references (§18.17.2.3): one rectangle read across a span of sheets, endpoints
/// included, in tab order.
/// </summary>
public sealed class ThreeDReferenceTests
{
    private static Workbook Book()
    {
        var book = Workbook.Create();
        for (int i = 1; i <= 4; i++)
        {
            Worksheet sheet = book.Sheets.Add($"Q{i}");
            sheet["A1"] = i * 10;
            sheet["A2"] = i;
        }

        return book;
    }

    [Fact]
    public void SumReadsEverySheetOfTheSpan()
    {
        Workbook book = Book();
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM(Q1:Q3!A1)");
        totals.Cell("B2").SetFormula("SUM(Q1:Q4!A1:A2)");
        totals.Cell("B3").SetFormula("COUNT(Q2:Q3!A1)");
        totals.Cell("B4").SetFormula("AVERAGE(Q1:Q4!A2)");
        book.Calculate();

        Assert.Equal(60.0, totals["B1"].GetNumber());
        Assert.Equal(110.0, totals["B2"].GetNumber());
        Assert.Equal(2.0, totals["B3"].GetNumber());
        Assert.Equal(2.5, totals["B4"].GetNumber());
    }

    [Fact]
    public void ReversedEndpointsReadTheSameSpan()
    {
        Workbook book = Book();
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM(Q3:Q1!A1)");
        book.Calculate();

        Assert.Equal(60.0, totals["B1"].GetNumber());
    }

    [Fact]
    public void QuotedSpanNamesParse()
    {
        var book = Workbook.Create();
        book.Sheets.Add("S 1")["A1"] = 5;
        book.Sheets.Add("S 2")["A1"] = 7;
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM('S 1:S 2'!A1)");
        book.Calculate();

        Assert.Equal(12.0, totals["B1"].GetNumber());
    }

    [Fact]
    public void AMissingEndpointIsAReferenceError()
    {
        Workbook book = Book();
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM(Q1:Q9!A1)");
        book.Calculate();

        Assert.Equal(ExcelError.Reference, totals["B1"].GetError());
    }

    [Fact]
    public void AChangeOnAMiddleSheetRecalculatesTheSpan()
    {
        Workbook book = Book();
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM(Q1:Q3!A1)");
        book.Calculate();
        Assert.Equal(60.0, totals["B1"].GetNumber());

        // The middle sheet is not named in the formula, but the span subscribes to it.
        book.Sheets["Q2"]!["A1"] = 200;
        CalculationSummary summary = book.Calculate();

        Assert.Equal(240.0, totals["B1"].GetNumber());
        Assert.True(summary.CalculatedCells >= 1);
    }

    [Fact]
    public void RenamingAnEndpointFollowsIntoTheFormula()
    {
        Workbook book = Book();
        Worksheet totals = book.Sheets.Add("Totals");
        totals.Cell("B1").SetFormula("SUM(Q1:Q3!A1)");

        book.Sheets["Q3"]!.Name = "Third Quarter";

        Assert.Equal("SUM('Q1:Third Quarter'!A1)", totals.Cell("B1").Formula);
        book.Calculate();
        Assert.Equal(60.0, totals["B1"].GetNumber());
    }
}
