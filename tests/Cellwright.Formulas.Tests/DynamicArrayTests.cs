using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The dynamic array functions. Without spilling, an array result is checked the way any array
/// formula has always been checked: through the function that consumes it.
/// </summary>
public sealed class DynamicArrayTests
{
    /// <summary>
    /// Three regions and their takings, deliberately out of order and with one repeat, so that
    /// sorting, filtering and de-duplicating all have something to do.
    /// </summary>
    private static Worksheet Sales()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        Row(sheet, 1, "North", 30);
        Row(sheet, 2, "South", 10);
        Row(sheet, 3, "East", 20);
        Row(sheet, 4, "North", 30);
        return sheet;
    }

    private static void Row(Worksheet sheet, int row, string region, double amount)
    {
        sheet[row, 1] = CellValue.Text(region);
        sheet[row, 2] = CellValue.Number(amount);
    }

    private static CellValue Eval(string formula) => Sales().Evaluate(formula, "E1");

    /// <summary>Rows travel whole: sorting on the takings carries each region along with them.</summary>
    [Theory]
    [InlineData("INDEX(SORT(A1:B4,2),1,1)", "South")]
    [InlineData("INDEX(SORT(A1:B4,2),2,1)", "East")]
    [InlineData("INDEX(SORT(A1:B4,2,-1),1,1)", "North")]
    [InlineData("INDEX(SORT(A1:B4,1),1,1)", "East")]
    public void SortMovesWholeRows(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Theory]
    [InlineData("INDEX(SORT(A1:B4,2),1,2)", 10.0)]
    [InlineData("INDEX(SORT(A1:B4,2),4,2)", 30.0)]
    [InlineData("SUM(SORT(B1:B4))", 90.0)]
    public void SortKeepsEveryValue(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Equal entries keep the order they were written in, which is the only way to tell that the
    /// sort is stable — and the only way a second sort key means anything.
    /// </summary>
    [Fact]
    public void EqualEntriesKeepTheirOrder()
    {
        Worksheet sheet = Sales();
        sheet[1, 3] = CellValue.Text("first");
        sheet[4, 3] = CellValue.Text("second");

        // East sorts ahead of both, and of the two North rows the one written first comes first.
        Assert.Equal("East", sheet.Evaluate("INDEX(SORT(A1:C4,1),1,1)", "E1").GetText());
        Assert.Equal("first", sheet.Evaluate("INDEX(SORT(A1:C4,1),2,3)", "E1").GetText());
        Assert.Equal("second", sheet.Evaluate("INDEX(SORT(A1:C4,1),3,3)", "E1").GetText());
    }

    /// <summary>The key of SORTBY need not be in the result, which is why it exists beside SORT.</summary>
    [Fact]
    public void SortByOrdersOnSomethingItDoesNotReturn()
    {
        Assert.Equal("South", Eval("INDEX(SORTBY(A1:A4,B1:B4),1,1)").GetText());
        Assert.Equal("North", Eval("INDEX(SORTBY(A1:A4,B1:B4,-1),1,1)").GetText());

        // Two keys: the takings first, then the region to settle the tie.
        Assert.Equal("East", Eval("INDEX(SORTBY(A1:A4,B1:B4,1,A1:A4,1),2,1)").GetText());
    }

    [Theory]
    [InlineData("SUM(FILTER(B1:B4,A1:A4=\"North\"))", 60.0)]
    [InlineData("SUM(FILTER(B1:B4,B1:B4>15))", 80.0)]
    [InlineData("COUNT(FILTER(B1:B4,B1:B4>15))", 3.0)]
    [InlineData("SUM(FILTER(B1:B4,A1:A4=\"West\",0))", 0.0)]
    public void FilterKeepsWhatThePredicateAdmits(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Nothing left is not an error in itself, but there is nothing to show, so the caller's own
    /// stand-in is used and <c>#CALC!</c> only appears when it offered none.
    /// </summary>
    [Fact]
    public void AnEmptyFilterFallsBackOrComplains()
    {
        Assert.Equal("none", Eval("FILTER(A1:A4,A1:A4=\"West\",\"none\")").GetText());
        Assert.Equal(ExcelError.Calc, Eval("FILTER(A1:A4,A1:A4=\"West\")").GetError());
    }

    [Fact]
    public void UniqueDropsRepeats()
    {
        Assert.Equal(3.0, Eval("COUNTA(UNIQUE(A1:A4))").GetNumber());
        Assert.Equal("North", Eval("INDEX(UNIQUE(A1:A4),1,1)").GetText());
        Assert.Equal("East", Eval("INDEX(UNIQUE(A1:A4),3,1)").GetText());

        // Asked for what appears exactly once, the repeated region drops out altogether.
        Assert.Equal(2.0, Eval("COUNTA(UNIQUE(A1:A4,FALSE,TRUE))").GetNumber());
        Assert.Equal("South", Eval("INDEX(UNIQUE(A1:A4,FALSE,TRUE),1,1)").GetText());
    }

    [Fact]
    public void TransposeTurnsTheArrayOnItsSide()
    {
        Assert.Equal("North", Eval("INDEX(TRANSPOSE(A1:B4),1,1)").GetText());
        Assert.Equal(30.0, Eval("INDEX(TRANSPOSE(A1:B4),2,1)").GetNumber());
        Assert.Equal(4.0, Eval("COLUMNS(TRANSPOSE(A1:B4))").GetNumber());
        Assert.Equal(2.0, Eval("ROWS(TRANSPOSE(A1:B4))").GetNumber());
    }

    [Theory]
    [InlineData("SUM(SEQUENCE(4))", 10.0)]
    [InlineData("SUM(SEQUENCE(2,3))", 21.0)]
    [InlineData("INDEX(SEQUENCE(2,3),2,1)", 4.0)]
    [InlineData("SUM(SEQUENCE(3,1,10,5))", 45.0)]
    [InlineData("ROWS(SEQUENCE(5,2))", 5.0)]
    public void SequenceCountsAcrossThenDown(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void AnEmptyShapeHasNothingToBuild()
    {
        Assert.Equal(ExcelError.Calc, Eval("SEQUENCE(0)").GetError());
        Assert.Equal(ExcelError.Value, Eval("SEQUENCE(-1)").GetError());
        Assert.Equal(ExcelError.Number, Eval("SEQUENCE(1000000,1000)").GetError());
    }

    [Fact]
    public void RandomArrayStaysBetweenItsBounds()
    {
        Assert.Equal(6.0, Eval("COUNT(RANDARRAY(2,3))").GetNumber());
        Assert.True(Eval("SUM(RANDARRAY(5,5,10,20))").GetNumber() is >= 250 and <= 500);
        Assert.Equal(0.0, Eval("SUM(RANDARRAY(4,1,7,7,TRUE))").GetNumber() % 7);
        Assert.Equal(ExcelError.Value, Eval("RANDARRAY(2,2,10,1)").GetError());
    }

    /// <summary>
    /// The <c>@</c> of a modern formula narrows a range to the one cell lining up with the
    /// formula. It is written as an operator but it is the SINGLE function underneath.
    /// </summary>
    [Fact]
    public void TheAtSignNarrowsARangeToOneValue()
    {
        Worksheet sheet = Sales();

        // Evaluated from row 2, the column narrows to the value on that row.
        Assert.Equal(10.0, sheet.Evaluate("@B1:B4", "E2").GetNumber());
        Assert.Equal(10.0, sheet.Evaluate("SINGLE(B1:B4)", "E2").GetNumber());
        Assert.Equal(30.0, sheet.Evaluate("@B1:B4", "E1").GetNumber());
    }

    /// <summary>An array result in a cell of its own shows its first value, there being no spill.</summary>
    [Fact]
    public void ACellHoldingAnArrayShowsItsFirstValue()
    {
        Assert.Equal("South", Eval("SORT(A1:B4,2)").GetText());
        Assert.Equal(1.0, Eval("SEQUENCE(3)").GetNumber());
    }

    [Theory]
    [InlineData("TEXTBEFORE(\"Red-Green-Blue\",\"-\")", "Red")]
    [InlineData("TEXTBEFORE(\"Red-Green-Blue\",\"-\",2)", "Red-Green")]
    [InlineData("TEXTBEFORE(\"Red-Green-Blue\",\"-\",-1)", "Red-Green")]
    [InlineData("TEXTBEFORE(\"Red-Green-Blue\",\"x\",1,1,0,\"missing\")", "missing")]
    [InlineData("TEXTBEFORE(\"Red-Green-Blue\",\"-\",3,0,1)", "Red-Green-Blue")]
    [InlineData("TEXTAFTER(\"Red-Green-Blue\",\"-\")", "Green-Blue")]
    [InlineData("TEXTAFTER(\"Red-Green-Blue\",\"-\",2)", "Blue")]
    [InlineData("TEXTAFTER(\"Red-Green-Blue\",\"-\",-1)", "Blue")]
    [InlineData("TEXTAFTER(\"Red-Green-Blue\",\"x\",1,0,0,\"none\")", "none")]
    [InlineData("INDEX(TEXTSPLIT(\"a,b;c,d\",\",\",\";\"),2,2)", "d")]
    [InlineData("COLUMNS(TEXTSPLIT(\"a,,b\",\",\",,TRUE))", "2")]
    public void TextBeforeAfterAndSplitExtractAndSliceText(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).ToString());
    }

    [Theory]
    [InlineData("ROWS(VSTACK(A1:B2,A3:B4))", 4.0)]
    [InlineData("COLUMNS(HSTACK(A1:A4,B1:B4))", 2.0)]
    [InlineData("SUM(TOCOL(B1:B4))", 90.0)]
    [InlineData("COLUMNS(TOROW(B1:B4))", 4.0)]
    [InlineData("ROWS(TAKE(A1:B4,2))", 2.0)]
    [InlineData("INDEX(TAKE(A1:B4,-1),1,1)", 0.0, "North")]
    [InlineData("ROWS(DROP(A1:B4,1))", 3.0)]
    [InlineData("INDEX(CHOOSEROWS(A1:B4,2,3),2,1)", 0.0, "East")]
    [InlineData("INDEX(CHOOSECOLS(A1:B4,2),3,1)", 20.0)]
    [InlineData("ROWS(WRAPROWS(B1:B4,2))", 2.0)]
    [InlineData("COLUMNS(WRAPCOLS(B1:B4,2))", 2.0)]
    [InlineData("INDEX(EXPAND(B1:B2,3,2,99),3,2)", 99.0)]
    [InlineData("LAMBDA(x,y,x+y)(12,30)", 42.0)]
    [InlineData("LET(scale,10,fn,LAMBDA(v,v*scale),fn(7))", 70.0)]
    [InlineData("LAMBDA(x,y,IF(ISOMITTED(y),x*2,x+y))(5,)", 10.0)]
    [InlineData("SUM(MAP(B1:B4,LAMBDA(x,x*2)))", 180.0)]
    [InlineData("REDUCE(10,B1:B4,LAMBDA(acc,x,acc+x))", 100.0)]
    [InlineData("INDEX(SCAN(0,B1:B4,LAMBDA(acc,x,acc+x)),4,1)", 90.0)]
    [InlineData("SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)))", 102.0)]
    [InlineData("SUM(BYROW(B1:B4,LAMBDA(r,SUM(r))))", 90.0)]
    [InlineData("INDEX(BYCOL(B1:B4,LAMBDA(c,MAX(c))),1,1)", 30.0)]
    public void ModernArrayShapeFunctionsReshapeAndCombineArrays(
        string formula, double expectedNumber, string? expectedText = null)
    {
        CellValue actual = Eval(formula);
        if (expectedText is not null)
            Assert.Equal(expectedText, actual.GetText());
        else
            Assert.Equal(expectedNumber, actual.GetNumber(), precision: 8);
    }

    [Fact]
    public void DefinedNameLambdaAndSpilledMakeArrayWorkOnWorksheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        book.Names.Add("Hypot", "LAMBDA(a,b,SQRT(a*a+b*b))");

        sheet.Cell("A1").SetFormula("Hypot(3,4)");
        sheet.Cell("B1").SetFormula("MAKEARRAY(2,2,LAMBDA(r,c,r+c))");
        sheet.Cell("D1").SetFormula("LAMBDA(x,x+1)");
        book.Calculate();

        Assert.Equal(5.0, sheet["A1"].GetNumber(), precision: 8);
        Assert.Equal(2.0, sheet["B1"].GetNumber(), precision: 8);
        Assert.Equal(3.0, sheet["C1"].GetNumber(), precision: 8);
        Assert.Equal(3.0, sheet["B2"].GetNumber(), precision: 8);
        Assert.Equal(4.0, sheet["C2"].GetNumber(), precision: 8);
        Assert.Equal(ExcelError.Calc, sheet["D1"].GetError());
    }
}
