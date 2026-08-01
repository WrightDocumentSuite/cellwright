using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Golden checks for the correlation, regression and single-sample statistics — nearly every
/// expected value is the worked example from Excel's own documentation.
/// </summary>
public sealed class StatisticalPackTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        return book.Sheets.Add("S").Evaluate(formula, "Z1");
    }

    [Theory]
    [InlineData("CORREL({3,2,4,5,6},{9,7,12,15,17})", 0.99705449)]
    [InlineData("PEARSON({9,7,5,3,1},{10,6,1,5,3})", 0.6993786)]
    [InlineData("RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.05795019)]
    [InlineData("COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    [InlineData("COVAR({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    [InlineData("COVARIANCE.S({2,4,8},{5,11,12})", 9.66666667)]
    [InlineData("SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.30555556)]
    [InlineData("INTERCEPT({2,3,9,1,8},{6,5,11,7,5})", 0.04838710)]
    [InlineData("STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 3.30571895)]
    [InlineData("FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})", 10.60725309)]
    [InlineData("FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})", 10.60725309)]
    public void PairedStatistics(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("TTEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)", 0.19601578)]
    [InlineData("T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)", 0.19601578)]
    [InlineData("TTEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,2)", 0.19199589)]
    [InlineData("FTEST({6,7,9,15,21},{20,28,31,38,40})", 0.64831785)]
    [InlineData("CHITEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})", 0.00030820)]
    [InlineData("ZTEST({3,6,7,8,6,5,4,2,1,9},4)", 0.09057419)]
    public void HypothesisTests(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("AVEDEV({4,5,6,7,5,4,3})", 1.02040816)]
    [InlineData("DEVSQ({4,5,8,7,11,4,3})", 48.0)]
    [InlineData("GEOMEAN({4,5,8,7,11,4,3})", 5.47698696)]
    [InlineData("HARMEAN({4,5,8,7,11,4,3})", 5.02837596)]
    [InlineData("TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)", 3.77777778)]
    [InlineData("KURT({3,4,5,2,3,4,5,6,4,7})", -0.15179963)]
    [InlineData("SKEW({3,4,5,2,3,4,5,6,4,7})", 0.35954307)]
    [InlineData("SKEW.P({3,4,5,2,3,4,5,6,4,7})", 0.30319331)]
    [InlineData("FISHER(0.75)", 0.97295507)]
    [InlineData("FISHERINV(0.972955)", 0.74999997)]
    [InlineData("PERCENTRANK({13,12,11,8,4,3,2,1,1,1},2)", 0.333)]
    [InlineData("PERCENTRANK({13,12,11,8,4,3,2,1,1,1},4)", 0.555)]
    [InlineData("PERCENTRANK({13,12,11,8,4,3,2,1,1,1},8,2)", 0.66)]
    [InlineData("PERCENTRANK({13,12,11,8,4,3,2,1,1,1},5)", 0.583)]
    [InlineData("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", 0.1)]
    [InlineData("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", 0.8)]
    [InlineData("RANK.AVG(94,{89,88,92,101,94,97,95})", 4.0)]
    [InlineData("RANK.AVG(2,{1,2,2,3})", 2.5)]
    [InlineData("STDEVA({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 27.46391572)]
    [InlineData("VARA({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 754.26666667)]
    [InlineData("STDEV.S({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 27.46391572)]
    [InlineData("VAR.P({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 678.84)]
    [InlineData("MODE.SNGL({5.6,4,4,3,2,4})", 4.0)]
    public void SingleSampleStatistics(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Fact]
    public void StdevaCountsWordsAsZeroes()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet["A2"] = 2;
        sheet["A3"] = "note";
        sheet["A4"] = true;

        // Values counted: 1, 2, 0, 1 — the plain STDEV would see only 1 and 2.
        Assert.Equal(0.816496581, sheet.Evaluate("STDEVA(A1:A4)", "Z1").GetNumber(), precision: 8);
        Assert.Equal(0.707106781, sheet.Evaluate("STDEV(A1:A4)", "Z1").GetNumber(), precision: 8);
    }

    [Fact]
    public void ModeMultReportsEveryMode()
    {
        Assert.Equal(2.0, Eval("INDEX(MODE.MULT({1,2,3,4,3,2,1,2,3}),1)").GetNumber());
        Assert.Equal(3.0, Eval("INDEX(MODE.MULT({1,2,3,4,3,2,1,2,3}),2)").GetNumber());
        Assert.Equal(ExcelError.NotAvailable, Eval("MODE.MULT({1,2,3})").GetError());
    }

    [Fact]
    public void TrendExtendsTheLine()
    {
        CellValue month13 = Eval(
            "INDEX(TREND({133890;135000;135790;137300;138130;139100;139900;141120;141890;143230;144000;145290},,{13}),1)");
        Assert.Equal(146171.5152, month13.GetNumber(), precision: 3);
    }

    [Fact]
    public void GrowthAndLogestFitTheExponential()
    {
        CellValue month17 = Eval(
            "INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},{17}),1)");
        Assert.Equal(320196.7184, month17.GetNumber(), precision: 3);

        CellValue m = Eval("INDEX(LOGEST({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16}),1,1)");
        CellValue b = Eval("INDEX(LOGEST({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16}),1,2)");
        Assert.Equal(1.46327582, m.GetNumber(), precision: 6);
        Assert.Equal(495.30477, b.GetNumber(), precision: 3);
    }

    [Fact]
    public void PairedFunctionsRefuseMismatchedShapes()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("CORREL({1,2,3},{1,2})").GetError());
        Assert.Equal(ExcelError.DivisionByZero, Eval("CORREL({1,1,1},{2,3,4})").GetError());
        Assert.Equal(ExcelError.Number, Eval("TTEST({1,2},{3,4},3,1)").GetError());
        Assert.Equal(ExcelError.Number, Eval("GEOMEAN({4,-1})").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("PERCENTRANK({1,2,3},9)").GetError());
    }
}
