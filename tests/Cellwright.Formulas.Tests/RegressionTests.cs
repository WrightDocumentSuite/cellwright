using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// <c>LINEST</c> against fits worked out by hand. The four-point set is deliberately imperfect,
/// since a perfect fit hides every mistake the statistics block could make.
/// </summary>
public sealed class RegressionTests
{
    private static Worksheet Data()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");

        // One variable, imperfect: the fit is y = 61/35 x + 64/35.
        Fill(sheet, "A", [2, 9, 5, 7]);
        Fill(sheet, "B", [0, 4, 2, 3]);

        // Two variables, exact: y = x1 + 2 x2 + 3.
        Fill(sheet, "D", [6, 11, 10, 17]);
        Fill(sheet, "E", [1, 2, 3, 4]);
        Fill(sheet, "F", [1, 3, 2, 5]);

        // Straight through the origin: y = 2x.
        Fill(sheet, "H", [2, 4, 6]);
        Fill(sheet, "I", [1, 2, 3]);

        for (int row = 1; row <= 4; row++)
            sheet[$"K{row}"] = "not a number";

        return sheet;
    }

    private static void Fill(Worksheet sheet, string column, double[] values)
    {
        for (int i = 0; i < values.Length; i++)
            sheet[$"{column}{i + 1}"] = CellValue.Number(values[i]);
    }

    private static CellValue Eval(string formula) => Data().Evaluate(formula, "Z1");

    [Theory]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4),1,1)", 1.74285714)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4),1,2)", 1.82857143)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),1,1)", 1.74285714)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),1,2)", 1.82857143)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),2,1)", 0.09897433)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),2,2)", 0.26649654)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),3,1)", 0.99359145)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),3,2)", 0.29277002)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),4,1)", 310.0833333)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),4,2)", 2.0)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),5,1)", 26.57857143)]
    [InlineData("INDEX(LINEST(A1:A4,B1:B4,TRUE,TRUE),5,2)", 0.17142857)]
    public void TheStatisticsBlockMatchesTheFitWorkedOutByHand(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// Excel lists the coefficients in reverse — the last variable first and the constant last —
    /// which is the opposite of how the columns are written.
    /// </summary>
    [Theory]
    [InlineData("INDEX(LINEST(D1:D4,E1:F4),1,1)", 2.0)]
    [InlineData("INDEX(LINEST(D1:D4,E1:F4),1,2)", 1.0)]
    [InlineData("INDEX(LINEST(D1:D4,E1:F4),1,3)", 3.0)]
    public void SeveralVariablesComeBackInReverse(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Left without a variable, the fit runs against the position of each observation, which is
    /// what makes <c>LINEST(y)</c> a trend over time.
    /// </summary>
    [Fact]
    public void WithoutAVariableTheFitRunsAgainstPosition()
    {
        Assert.Equal(1.1, Eval("INDEX(LINEST(A1:A4),1,1)").GetNumber(), precision: 8);
        Assert.Equal(3.0, Eval("INDEX(LINEST(A1:A4),1,2)").GetNumber(), precision: 8);
    }

    /// <summary>A fit told to pass through the origin still reports a constant, as a zero.</summary>
    [Fact]
    public void AFitThroughTheOriginReportsNoConstant()
    {
        Assert.Equal(2.0, Eval("INDEX(LINEST(H1:H3,I1:I3,FALSE),1,1)").GetNumber(), precision: 8);
        Assert.Equal(0.0, Eval("INDEX(LINEST(H1:H3,I1:I3,FALSE),1,2)").GetNumber(), precision: 8);
        Assert.True(Eval("ISNA(INDEX(LINEST(H1:H3,I1:I3,FALSE,TRUE),2,2))").GetBoolean());
    }

    /// <summary>
    /// The block is always five rows by one more than the number of variables, so the corner
    /// beyond the two statistics each row carries has to say something. Excel says <c>#N/A</c>.
    /// </summary>
    [Fact]
    public void TheCornersOfTheBlockAreNotAvailable()
    {
        Assert.True(Eval("ISNA(INDEX(LINEST(D1:D4,E1:F4,TRUE,TRUE),3,3))").GetBoolean());
        Assert.True(Eval("ISNA(INDEX(LINEST(D1:D4,E1:F4,TRUE,TRUE),5,3))").GetBoolean());
    }

    /// <summary>A cell holding the fit shows its first number, there being nowhere to spill.</summary>
    [Fact]
    public void ACellHoldingTheFitShowsTheFirstCoefficient()
    {
        Assert.Equal(1.74285714, Eval("LINEST(A1:A4,B1:B4)").GetNumber(), precision: 6);
    }

    /// <summary>
    /// Observations and variables that do not line up are a mistake in the ranges rather than
    /// in the data, and Excel says so with a different error than it uses for unusable values.
    /// </summary>
    [Fact]
    public void MismatchedShapesAreRefused()
    {
        Assert.Equal(ExcelError.Reference, Eval("LINEST(A1:A4,B1:B3)").GetError());
        Assert.Equal(ExcelError.Reference, Eval("LINEST(A1:B4,B1:B4)").GetError());
        Assert.Equal(ExcelError.Value, Eval("LINEST(A1:A4,K1:K4)").GetError());
    }
}
