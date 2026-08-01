using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Golden checks for the securities catalog: coupon calendars, bond prices and yields, bills,
/// accrued interest, depreciations and the odd-period bonds. Expected values are the worked
/// examples of Excel's documentation.
/// </summary>
public sealed class SecuritiesPackTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        return book.Sheets.Add("S").Evaluate(formula, "Z1");
    }

    [Theory]
    [InlineData("EFFECT(0.0525,4)", 0.05354266737)]
    [InlineData("NOMINAL(0.053543,4)", 0.05250032)]
    [InlineData("CUMIPMT(0.09/12,360,125000,13,24,0)", -11135.23213)]
    [InlineData("CUMIPMT(0.09/12,360,125000,1,1,0)", -937.5)]
    [InlineData("CUMPRINC(0.09/12,360,125000,13,24,0)", -934.10712)]
    [InlineData("CUMPRINC(0.09/12,360,125000,1,1,0)", -68.27827)]
    [InlineData("ISPMT(0.1/12,1,3*12,8000000)", -64814.81481)]
    [InlineData("FVSCHEDULE(1,{0.09,0.11,0.1})", 1.33089)]
    [InlineData("DOLLARDE(1.02,16)", 1.125)]
    [InlineData("DOLLARDE(1.1,32)", 1.3125)]
    [InlineData("DOLLARFR(1.125,16)", 1.02)]
    [InlineData("DOLLARFR(1.125,32)", 1.04)]
    [InlineData("VDB(2400,300,10,0,1)", 480.0)]
    [InlineData("VDB(2400,300,10,1,2)", 384.0)]
    [InlineData("VDB(2400,300,10,0,0.875,1.5)", 315.0)]
    public void RatesAndSchedules(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 5)
;
    }

    [Theory]
    [InlineData("COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)", 71.0)]
    [InlineData("COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,1)", 181.0)]
    [InlineData("COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)", 110.0)]
    [InlineData("COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)", 4.0)]
    public void CouponCounts(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber());
    }

    [Fact]
    public void CouponDatesFollowTheCalendar()
    {
        Assert.Equal(new DateTime(2011, 5, 15), Eval("COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)").GetDateTime());
        Assert.Equal(new DateTime(2010, 11, 15), Eval("COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)").GetDateTime());

        // A maturity on a month's last day pins every coupon to month ends.
        Assert.Equal(new DateTime(2008, 2, 29), Eval("COUPPCD(DATE(2008,3,15),DATE(2008,8,31),2,0)").GetDateTime());
    }

    [Theory]
    [InlineData("PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", 94.63436162)]
    [InlineData("YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)", 0.065)]
    [InlineData("DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", 5.99377496)]
    [InlineData("MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", 5.73566982)]
    [InlineData("DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)", 0.05242021)]
    [InlineData("PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)", 99.79583333)]
    [InlineData("YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)", 0.05282257)]
    [InlineData("INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)", 0.05768)]
    [InlineData("RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)", 1014584.6544071)]
    [InlineData("PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)", 99.98449888)]
    [InlineData("YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)", 0.06095434)]
    public void BondsAndDiscountedPaper(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", 0.09415149)]
    [InlineData("TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)", 98.45)]
    [InlineData("TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)", 0.09141696)]
    public void TreasuryBills(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)", 16.66666667)]
    [InlineData("ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE)", 15.55555556)]
    [InlineData("ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)", 20.54794521)]
    public void AccruedInterest(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)", 360.0)]
    [InlineData("AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)", 776.0)]
    public void FrenchDepreciation(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 5);
    }

    [Theory]
    [InlineData("ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)", 99.87828631, 6)]
    [InlineData("ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)", 0.04519223, 6)]
    // Excel's documentation rounds these two to 113.60 and 7.72%; the four-decimal agreement
    // is what the published examples pin down.
    [InlineData("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)", 113.5977, 3)]
    [InlineData("ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)", 0.07725, 4)]
    public void OddPeriodBonds(string formula, double expected, int precision)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision);
    }

    [Fact]
    public void ConnectionBackedFunctionsAnswerHonestly()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("CUBEVALUE(\"model\",\"[Measures].[x]\")").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("RTD(\"prog\",\"srv\",\"topic\")").GetError());
        Assert.Equal(ExcelError.Reference, Eval("GETPIVOTDATA(\"Sales\",A3)").GetError());
    }

    [Fact]
    public void BondArgumentsAreValidated()
    {
        Assert.Equal(ExcelError.Number, Eval("PRICE(DATE(2018,1,1),DATE(2008,1,1),0.05,0.06,100,2)").GetError());
        Assert.Equal(ExcelError.Number, Eval("PRICE(DATE(2008,1,1),DATE(2018,1,1),0.05,0.06,100,3)").GetError());
        Assert.Equal(ExcelError.Number, Eval("COUPNUM(DATE(2008,1,1),DATE(2018,1,1),2,7)").GetError());
        Assert.Equal(ExcelError.Number, Eval("TBILLPRICE(DATE(2008,1,1),DATE(2010,1,1),0.09)").GetError());
    }
}
