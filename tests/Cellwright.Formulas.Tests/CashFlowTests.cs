using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The dated cash flow functions and the four ways of writing an asset down, against the
/// worked examples Excel's documentation publishes for each.
/// </summary>
public sealed class CashFlowTests
{
    /// <summary>
    /// The schedule from Excel's own XIRR example: an outlay and four uneven receipts, on dates
    /// that are nothing like evenly spaced.
    /// </summary>
    private static Worksheet Ledger()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        Entry(sheet, 1, -10000, "2008-01-01");
        Entry(sheet, 2, 2750, "2008-03-01");
        Entry(sheet, 3, 4250, "2008-10-30");
        Entry(sheet, 4, 3250, "2009-02-15");
        Entry(sheet, 5, 2750, "2009-04-01");

        // An even series for the modified rate of return, which counts periods rather than days.
        double[] flows = [-120000, 39000, 30000, 21000, 37000, 46000];
        for (int i = 0; i < flows.Length; i++)
            sheet[i + 1, 4] = CellValue.Number(flows[i]);

        return sheet;
    }

    private static void Entry(Worksheet sheet, int row, double amount, string date)
    {
        sheet[row, 1] = CellValue.Number(amount);
        sheet[row, 2] = CellValue.DateTime(DateTime.Parse(date, System.Globalization.CultureInfo.InvariantCulture));
    }

    private static CellValue Eval(string formula) => Ledger().Evaluate(formula, "F1");

    [Theory]
    [InlineData("XIRR(A1:A5,B1:B5)", 0.373362535)]
    [InlineData("XIRR(A1:A5,B1:B5,0.1)", 0.373362535)]
    [InlineData("MIRR(D1:D6,0.1,0.12)", 0.126094130)]
    [InlineData("MIRR(D1:D4,0.1,0.12)", -0.048044655)]
    [InlineData("MIRR(D1:D6,0.1,0.14)", 0.134759111)]
    public void TheDatedFlowsAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>Excel publishes this one to four decimals, so four is what is asserted.</summary>
    [Fact]
    public void DatedDiscountingAgreesWithExcel()
    {
        Assert.Equal(2086.6476, Eval("XNPV(0.09,A1:A5,B1:B5)").GetNumber(), precision: 4);
    }

    /// <summary>
    /// The rate <c>XIRR</c> finds is the one at which <c>XNPV</c> comes to nothing, which is
    /// what it means, and checking it that way needs no published digits at all.
    /// </summary>
    [Fact]
    public void TheRateFoundIsTheRateThatBalances()
    {
        double rate = Eval("XIRR(A1:A5,B1:B5)").GetNumber();
        string text = rate.ToString(System.Globalization.CultureInfo.InvariantCulture);

        Assert.Equal(0.0, Eval($"XNPV({text},A1:A5,B1:B5)").GetNumber(), precision: 6);
    }

    [Theory]
    [InlineData("SLN(30000,7500,10)", 2250.0)]
    [InlineData("SYD(30000,7500,10,1)", 4090.909091)]
    [InlineData("SYD(30000,7500,10,10)", 409.090909)]
    [InlineData("DDB(2400,300,10,1)", 480.0)]
    [InlineData("DDB(2400,300,10*12,1)", 40.0)]
    [InlineData("DDB(2400,300,10*365,1)", 1.315068)]
    [InlineData("DDB(2400,300,10,2,1.5)", 306.0)]
    [InlineData("DDB(2400,300,10,10)", 22.122547)]
    public void TheDepreciationsAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// Excel's own <c>DB</c> table: a million written down to a hundred thousand over six years,
    /// bought in the seventh month, so the first and last years are part years.
    /// </summary>
    [Theory]
    [InlineData(1, 186083.333333)]
    [InlineData(2, 259639.416667)]
    [InlineData(3, 176814.442750)]
    [InlineData(4, 120410.635513)]
    [InlineData(5, 81999.642784)]
    [InlineData(6, 55841.756736)]
    [InlineData(7, 15845.098474)]
    public void TheFixedDecliningBalanceFollowsExcelsTable(int period, double expected)
    {
        Assert.Equal(expected, Eval($"DB(1000000,100000,6,{period},7)").GetNumber(), precision: 5);
    }

    /// <summary>
    /// Written down over its whole life, the charges add up to what the asset lost. That holds
    /// for every method, and it is the one property none of them may break.
    /// </summary>
    [Fact]
    public void TheChargesAddUpToWhatWasLost()
    {
        double straight = 0;
        double years = 0;
        double declining = 0;
        for (int period = 1; period <= 10; period++)
        {
            straight += Eval("SLN(30000,7500,10)").GetNumber();
            years += Eval($"SYD(30000,7500,10,{period})").GetNumber();
            declining += Eval($"DDB(30000,7500,10,{period})").GetNumber();
        }

        Assert.Equal(22500.0, straight, precision: 6);
        Assert.Equal(22500.0, years, precision: 6);
        Assert.Equal(22500.0, declining, precision: 6);
    }

    [Theory]
    [InlineData("SYD(30000,7500,10,11)")]
    [InlineData("SYD(30000,7500,0,1)")]
    [InlineData("DDB(2400,300,10,11)")]
    [InlineData("DDB(2400,300,0,1)")]
    [InlineData("DDB(2400,300,10,1,0)")]
    [InlineData("DB(1000000,100000,6,8,7)")]
    [InlineData("DB(1000000,100000,6,1,13)")]
    [InlineData("DB(1000000,100000,0,1,7)")]
    [InlineData("XNPV(-1,A1:A5,B1:B5)")]
    [InlineData("XIRR(A2:A5,B2:B5)")]
    public void ImpossibleArgumentsAreRefused(string formula)
    {
        Assert.Equal(ExcelError.Number, Eval(formula).GetError());
    }

    [Fact]
    public void ALifeOfNothingHasNothingToSpread()
    {
        Assert.Equal(ExcelError.DivisionByZero, Eval("SLN(30000,7500,0)").GetError());
        Assert.Equal(ExcelError.DivisionByZero, Eval("MIRR(D2:D6,0.1,0.12)").GetError());
    }
}
