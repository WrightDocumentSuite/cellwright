using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The annuity family and the discounted cash flow pair, against values Excel's own
/// documentation states for the same arguments.
/// </summary>
public sealed class FinancialFunctionTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = -70000;
        sheet["A2"] = 12000;
        sheet["A3"] = 15000;
        sheet["A4"] = 18000;
        sheet["A5"] = 21000;
        sheet["A6"] = 26000;
        return sheet.Evaluate(formula, "Z1");
    }

    [Theory]
    [InlineData("PMT(0.08/12,60,30000)", -608.29)]
    [InlineData("PMT(0.1/12,2*12,2000)", -92.29)]
    [InlineData("PV(0.08/12,12*20,500)", -59777.15)]
    [InlineData("FV(0.06/12,10,-200,-500,1)", 2581.40)]
    [InlineData("IPMT(0.1/12,1,3*12,8000)", -66.67)]
    [InlineData("PPMT(0.1/12,1,2*12,2000)", -75.62)]
    public void TheAnnuityFamilyAgreesWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 2);
    }

    [Theory]
    [InlineData("NPER(0.12/12,-100,-1000,10000,1)", 59.6738657)]
    [InlineData("RATE(4*12,-200,8000)", 0.00770147)]
    [InlineData("NPV(0.1,-10000,3000,4200,6800)", 1188.443412)]
    [InlineData("IRR(A1:A6)", 0.0866309)]
    [InlineData("IRR(A1:A5)", -0.0212448)]
    public void TheSearchedRatesAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// A guess only says where to start looking. A venture that turns around late has a steeply
    /// negative rate over its first years, and the search has to reach it from the default start
    /// as readily as from a guess nearby.
    /// </summary>
    [Theory]
    [InlineData("IRR(A1:A6,-0.9)", 0.0866309)]
    [InlineData("IRR(A1:A6,10)", 0.0866309)]
    [InlineData("IRR(A1:A3,-0.1)", -0.4435069)]
    [InlineData("IRR(A1:A3)", -0.4435069)]
    public void TheGuessOnlyChoosesWhereToStartLooking(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// A rate of zero is not a special case anybody writes on purpose, but it falls out of a
    /// spreadsheet often enough — an interest-free plan, a blank cell — and the compounding
    /// formulas divide by it, so each function needs its own straight-line answer.
    /// </summary>
    [Theory]
    [InlineData("PMT(0,10,1000)", -100.0)]
    [InlineData("PV(0,10,-100)", 1000.0)]
    [InlineData("FV(0,10,-100)", 1000.0)]
    [InlineData("NPER(0,-100,1000)", 10.0)]
    public void ARateOfZeroLeavesTheMoneyPilingUp(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Paying at the start of a period rather than the end earns a period less interest, and
    /// the very first payment earns none at all.
    /// </summary>
    [Fact]
    public void APaymentAtTheStartOfThePeriodCostsLess()
    {
        double atEnd = Eval("PMT(0.08/12,60,30000,0,0)").GetNumber();
        double atStart = Eval("PMT(0.08/12,60,30000,0,1)").GetNumber();
        Assert.True(Math.Abs(atStart) < Math.Abs(atEnd));

        Assert.Equal(0.0, Eval("IPMT(0.1/12,1,3*12,8000,0,1)").GetNumber(), precision: 10);
        Assert.Equal(Eval("PMT(0.1/12,3*12,8000,0,1)").GetNumber(), Eval("PPMT(0.1/12,1,3*12,8000,0,1)").GetNumber(), precision: 8);
    }

    /// <summary>Interest and principal are the two halves of one payment, period by period.</summary>
    [Fact]
    public void InterestAndPrincipalAddUpToThePayment()
    {
        double payment = Eval("PMT(0.09/12,36,15000)").GetNumber();
        for (int period = 1; period <= 36; period++)
        {
            double interest = Eval($"IPMT(0.09/12,{period},36,15000)").GetNumber();
            double principal = Eval($"PPMT(0.09/12,{period},36,15000)").GetNumber();
            Assert.Equal(payment, interest + principal, precision: 8);
        }
    }

    /// <summary>
    /// <c>NPV</c> discounts from the first period and <c>IRR</c> from the zeroth, so an outlay
    /// belongs inside the range of one and in front of the other. Stated that way the two agree.
    /// </summary>
    [Fact]
    public void NpvAtTheInternalRateComesToNothing()
    {
        double rate = Eval("IRR(A1:A6)").GetNumber();
        double present = Eval($"-70000+NPV({rate.ToString(System.Globalization.CultureInfo.InvariantCulture)},A2:A6)").GetNumber();
        Assert.Equal(0.0, present, precision: 6);
    }

    /// <summary>Values a range cannot offer a number for are passed over, not counted as zero.</summary>
    [Fact]
    public void NpvPassesOverWhatARangeCannotOffer()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 100;
        sheet["A2"] = "note";
        sheet["A4"] = 200;

        Assert.Equal(
            sheet.Evaluate("NPV(0.1,100,200)", "Z1").GetNumber(),
            sheet.Evaluate("NPV(0.1,A1:A4)", "Z1").GetNumber(),
            precision: 8);
    }

    [Theory]
    [InlineData("IPMT(0.1/12,0,12,1000)")]
    [InlineData("IPMT(0.1/12,13,12,1000)")]
    [InlineData("PPMT(0.1/12,13,12,1000)")]
    [InlineData("NPER(0,0,1000)")]
    [InlineData("NPER(0.1,-100,1000,-100000)")]
    [InlineData("IRR(A2:A6)")]
    [InlineData("RATE(0,-200,8000)")]
    public void ImpossibleArgumentsAreRefused(string formula)
    {
        Assert.Equal(ExcelError.Number, Eval(formula).GetError());
    }

    [Fact]
    public void ARateOfMinusOneHasNothingToDiscountBy()
    {
        Assert.Equal(ExcelError.DivisionByZero, Eval("NPV(-1,100,200)").GetError());
    }
}
