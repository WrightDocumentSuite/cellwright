using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Golden checks for the catalog additions: hyperbolics, combinatorics, matrices, LOOKUP,
/// day-count fractions, currency text and Roman numerals — values verified in Excel, most of
/// them straight from its documentation examples.
/// </summary>
public sealed class ExtendedFunctionTests
{
    private static Worksheet Sheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 4.14;
        sheet["A2"] = 4.19;
        sheet["A3"] = 5.17;
        sheet["A4"] = 5.77;
        sheet["A5"] = 6.39;
        sheet["B1"] = "red";
        sheet["B2"] = "orange";
        sheet["B3"] = "yellow";
        sheet["B4"] = "green";
        sheet["B5"] = "blue";
        return sheet;
    }

    private static CellValue Eval(string formula) => Sheet().Evaluate(formula, "Z1");

    [Theory]
    [InlineData("SINH(1)", 1.1752011936)]
    [InlineData("COSH(1)", 1.5430806348)]
    [InlineData("TANH(0.5)", 0.4621171573)]
    [InlineData("ASINH(1)", 0.8813735870)]
    [InlineData("ACOSH(2)", 1.3169578969)]
    [InlineData("ATANH(0.5)", 0.5493061443)]
    [InlineData("SEC(0)", 1.0)]
    [InlineData("ACOT(1)", 0.7853981634)]
    public void HyperbolicsAndFriends(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void HyperbolicsRefuseTheirDomainEdges()
    {
        Assert.Equal(ExcelError.Number, Eval("ACOSH(0.5)").GetError());
        Assert.Equal(ExcelError.Number, Eval("ATANH(1)").GetError());
    }

    [Theory]
    [InlineData("COMBIN(8,2)", 28.0)]
    [InlineData("COMBIN(49,6)", 13983816.0)]
    [InlineData("PERMUT(8,2)", 56.0)]
    [InlineData("PERMUT(100,3)", 970200.0)]
    [InlineData("MROUND(10,3)", 9.0)]
    [InlineData("MROUND(-10,-3)", -9.0)]
    [InlineData("MROUND(1.3,0.2)", 1.4)]
    [InlineData("SQRTPI(1)", 1.7724538509)]
    [InlineData("SQRTPI(2)", 2.5066282746)]
    [InlineData("FACTDOUBLE(6)", 48.0)]
    [InlineData("FACTDOUBLE(7)", 105.0)]
    [InlineData("FACTDOUBLE(0)", 1.0)]
    [InlineData("MULTINOMIAL(2,3,4)", 1260.0)]
    [InlineData("SUMSQ(3,4)", 25.0)]
    [InlineData("SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", -55.0)]
    [InlineData("SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 521.0)]
    [InlineData("SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 79.0)]
    public void CombinatoricsAndPairedSums(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void SeriesSumApproximatesTheCosine()
    {
        CellValue value = Eval("SERIESSUM(PI()/4,0,2,{1,-0.5,0.0416666666666667,-0.00138888888888889})");
        Assert.Equal(0.7071032148, value.GetNumber(), precision: 8);
    }

    [Fact]
    public void CombinatoricsRefuseNegativesAndCrossedSigns()
    {
        Assert.Equal(ExcelError.Number, Eval("COMBIN(2,3)").GetError());
        Assert.Equal(ExcelError.Number, Eval("PERMUT(3,-1)").GetError());
        Assert.Equal(ExcelError.Number, Eval("MROUND(5,-2)").GetError());
        Assert.Equal(ExcelError.Number, Eval("FACTDOUBLE(-2)").GetError());
        Assert.Equal(ExcelError.Number, Eval("SQRTPI(-1)").GetError());
    }

    [Fact]
    public void PairedSumsWantMatchingShapes()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("SUMXMY2({1,2,3},{1,2})").GetError());
    }

    [Theory]
    [InlineData("INDEX(MMULT({1,2;3,4},{5,6;7,8}),1,1)", 19.0)]
    [InlineData("INDEX(MMULT({1,2;3,4},{5,6;7,8}),2,2)", 50.0)]
    [InlineData("SUM(MMULT({1,2;3,4},{5,6;7,8}))", 134.0)]
    [InlineData("INDEX(MINVERSE({4,-1;2,0}),1,2)", 0.5)]
    [InlineData("INDEX(MINVERSE({4,-1;2,0}),2,1)", -1.0)]
    [InlineData("MDETERM({3,6,1;1,1,0;3,10,2})", 1.0)]
    [InlineData("MDETERM({1,2;2,4})", 0.0)]
    [InlineData("MDETERM({6})", 6.0)]
    public void MatrixFunctions(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void MatrixFunctionsRefuseBadShapesAndWords()
    {
        Assert.Equal(ExcelError.Value, Eval("MMULT({1,2},{1,2})").GetError());
        Assert.Equal(ExcelError.Value, Eval("MDETERM({1,2,3;4,5,6})").GetError());
        Assert.Equal(ExcelError.Value, Eval("MDETERM({1,\"x\";2,3})").GetError());
        Assert.Equal(ExcelError.Number, Eval("MINVERSE({1,2;2,4})").GetError());
    }

    [Fact]
    public void LookupVectorFormAnswersFromTheResultVector()
    {
        Worksheet sheet = Sheet();
        Assert.Equal("orange", sheet.Evaluate("LOOKUP(4.19,A1:A5,B1:B5)", "Z1").GetText());
        Assert.Equal("yellow", sheet.Evaluate("LOOKUP(5.75,A1:A5,B1:B5)", "Z1").GetText());
        Assert.Equal("blue", sheet.Evaluate("LOOKUP(7.66,A1:A5,B1:B5)", "Z1").GetText());
        Assert.Equal(ExcelError.NotAvailable, sheet.Evaluate("LOOKUP(0,A1:A5,B1:B5)", "Z1").GetError());
    }

    [Fact]
    public void LookupArrayFormPicksTheAxisByShape()
    {
        // Wider than tall: search the first row, answer from the last.
        Assert.Equal(3.0, Eval("LOOKUP(\"c\",{\"a\",\"b\",\"c\",\"d\";1,2,3,4})").GetNumber());

        // Tall or square: search the first column, answer from the last.
        Assert.Equal(20.0, Eval("LOOKUP(2,{1,10;2,20;3,30})").GetNumber());
    }

    [Fact]
    public void AreasCountsTheRectangles()
    {
        Assert.Equal(1.0, Eval("AREAS(A1:B3)").GetNumber());
        Assert.Equal(1.0, Eval("AREAS(A1)").GetNumber());
        Assert.Equal(3.0, Eval("AREAS((A1:A3,B1,C2:D4))").GetNumber());
        Assert.Equal(ExcelError.Value, Eval("AREAS(5)").GetError());
    }

    [Theory]
    [InlineData("YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))", 0.5805555556)]
    [InlineData("YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 0.5765027322)]
    [InlineData("YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)", 0.5861111111)]
    [InlineData("YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)", 0.5780821918)]
    [InlineData("YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)", 0.5805555556)]
    [InlineData("YEARFRAC(DATE(2020,2,29),DATE(2020,8,31),0)", 0.5)]
    [InlineData("YEARFRAC(DATE(2019,12,31),DATE(2020,1,2),1)", 0.0054794521)]
    [InlineData("YEARFRAC(DATE(2024,3,1),DATE(2024,3,1),1)", 0.0)]
    public void YearFracFollowsItsBasis(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void YearFracRefusesAnUnknownBasis()
    {
        Assert.Equal(ExcelError.Number, Eval("YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),5)").GetError());
    }

    [Theory]
    [InlineData("DOLLAR(1234.567,2)", "$1,234.57")]
    [InlineData("DOLLAR(1234.567,-2)", "$1,200")]
    [InlineData("DOLLAR(-1234.567,-2)", "($1,200)")]
    [InlineData("DOLLAR(-0.123,4)", "($0.1230)")]
    [InlineData("DOLLAR(99.888)", "$99.89")]
    [InlineData("USDOLLAR(2.5,1)", "$2.5")]
    public void DollarRendersCurrencyText(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Theory]
    [InlineData("ROMAN(499,0)", "CDXCIX")]
    [InlineData("ROMAN(499,1)", "LDVLIV")]
    [InlineData("ROMAN(499,2)", "XDIX")]
    [InlineData("ROMAN(499,3)", "VDIV")]
    [InlineData("ROMAN(499,4)", "ID")]
    [InlineData("ROMAN(2013)", "MMXIII")]
    [InlineData("ROMAN(1999,4)", "MIM")]
    [InlineData("ROMAN(0)", "")]
    [InlineData("ROMAN(499,FALSE)", "ID")]
    public void RomanWritesEachForm(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Fact]
    public void RomanAndArabicRoundTrip()
    {
        Assert.Equal(ExcelError.Value, Eval("ROMAN(-1)").GetError());
        Assert.Equal(ExcelError.Value, Eval("ROMAN(4000)").GetError());
        Assert.Equal(1999.0, Eval("ARABIC(\"MIM\")").GetNumber());
        Assert.Equal(1912.0, Eval("ARABIC(\"mcmxii\")").GetNumber());
        Assert.Equal(0.0, Eval("ARABIC(\"\")").GetNumber());
        Assert.Equal(ExcelError.Value, Eval("ARABIC(\"MB\")").GetError());
    }

    [Fact]
    public void CellReportsOnItsReference()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["B2"] = 42.5;
        sheet.Cell("B2").Style(s => s.NumberFormat("#,##0.00"));
        sheet["B3"] = "text";
        sheet.Cell("B4").SetValue(0.25).Style(s => s.NumberFormat("0.0%"));
        sheet.Cell("B5").SetValue(new DateTime(2026, 7, 28)).Style(s => s.NumberFormat("yyyy-mm-dd"));

        Assert.Equal("$B$2", sheet.Evaluate("CELL(\"address\",B2)", "Z1").GetText());
        Assert.Equal(2.0, sheet.Evaluate("CELL(\"col\",B2)", "Z1").GetNumber());
        Assert.Equal(2.0, sheet.Evaluate("CELL(\"row\",B2)", "Z1").GetNumber());
        Assert.Equal(42.5, sheet.Evaluate("CELL(\"contents\",B2)", "Z1").GetNumber());
        Assert.Equal("v", sheet.Evaluate("CELL(\"type\",B2)", "Z1").GetText());
        Assert.Equal("l", sheet.Evaluate("CELL(\"type\",B3)", "Z1").GetText());
        Assert.Equal("b", sheet.Evaluate("CELL(\"type\",B9)", "Z1").GetText());
        Assert.Equal(",2", sheet.Evaluate("CELL(\"format\",B2)", "Z1").GetText());
        Assert.Equal("P1", sheet.Evaluate("CELL(\"format\",B4)", "Z1").GetText());
        Assert.Equal("D4", sheet.Evaluate("CELL(\"format\",B5)", "Z1").GetText());
        Assert.Equal("G", sheet.Evaluate("CELL(\"format\",B3)", "Z1").GetText());
        Assert.Equal(1.0, sheet.Evaluate("CELL(\"protect\",B2)", "Z1").GetNumber());
        Assert.Equal(ExcelError.Value, sheet.Evaluate("CELL(\"nonsense\",B2)", "Z1").GetError());
    }

    [Fact]
    public void InfoAnswersWhatTheModelKnows()
    {
        var book = Workbook.Create();
        book.Sheets.Add("A");
        Worksheet sheet = book.Sheets.Add("B");

        Assert.Equal(2.0, sheet.Evaluate("INFO(\"numfile\")", "Z1").GetNumber());
        Assert.Equal("Manual", sheet.Evaluate("INFO(\"recalc\")", "Z1").GetText());
        Assert.Equal(ExcelError.NotAvailable, sheet.Evaluate("INFO(\"directory\")", "Z1").GetError());
        Assert.Equal(ExcelError.Value, sheet.Evaluate("INFO(\"nonsense\")", "Z1").GetError());
    }

    [Fact]
    public void ByteTwinsAndSpecNamesResolve()
    {
        Assert.Equal("ab", Eval("LEFTB(\"abc\",2)").GetText());
        Assert.Equal(3.0, Eval("LENB(\"abc\")").GetNumber());
        Assert.Equal(5.0, Eval("ECMA.CEILING(4.2,1)").GetNumber());
    }

    [Fact]
    public void FrequencyCountsIntoTheBins()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("F");
        double[] scores = [79, 85, 78, 85, 50, 81, 95, 88, 97];
        for (int i = 0; i < scores.Length; i++)
            sheet[$"A{i + 1}"] = scores[i];
        sheet["B1"] = 70;
        sheet["B2"] = 79;
        sheet["B3"] = 89;

        Assert.Equal(1.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,B1:B3),1)", "Z1").GetNumber());
        Assert.Equal(2.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,B1:B3),2)", "Z1").GetNumber());
        Assert.Equal(4.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,B1:B3),3)", "Z1").GetNumber());
        Assert.Equal(2.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,B1:B3),4)", "Z1").GetNumber());
        Assert.Equal(9.0, sheet.Evaluate("SUM(FREQUENCY(A1:A9,B1:B3))", "Z1").GetNumber());

        // Unsorted bins report their counts in the bins' own order.
        sheet["C1"] = 89;
        sheet["C2"] = 70;
        sheet["C3"] = 79;
        Assert.Equal(4.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,C1:C3),1)", "Z1").GetNumber());
        Assert.Equal(1.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,C1:C3),2)", "Z1").GetNumber());
        Assert.Equal(2.0, sheet.Evaluate("INDEX(FREQUENCY(A1:A9,C1:C3),3)", "Z1").GetNumber());
    }
}
