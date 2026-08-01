using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Golden checks for the engineering catalog: bases, unit conversions, complex arithmetic and
/// the Bessel families — expected values from Excel's documentation examples and the standard
/// tables of the functions themselves.
/// </summary>
public sealed class EngineeringPackTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        return book.Sheets.Add("S").Evaluate(formula, "Z1");
    }

    [Theory]
    [InlineData("BIN2DEC(1100100)", 100.0)]
    [InlineData("BIN2DEC(1111111111)", -1.0)]
    [InlineData("HEX2DEC(\"A5\")", 165.0)]
    [InlineData("HEX2DEC(\"FFFFFFFF5B\")", -165.0)]
    [InlineData("HEX2DEC(\"3DA408B9\")", 1034160313.0)]
    [InlineData("OCT2DEC(54)", 44.0)]
    [InlineData("OCT2DEC(7777777533)", -165.0)]
    public void BasesToDecimal(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber());
    }

    [Theory]
    [InlineData("DEC2BIN(9,4)", "1001")]
    [InlineData("DEC2BIN(-100)", "1110011100")]
    [InlineData("DEC2HEX(100,4)", "0064")]
    [InlineData("DEC2HEX(-54)", "FFFFFFFFCA")]
    [InlineData("DEC2HEX(28)", "1C")]
    [InlineData("DEC2OCT(58,3)", "072")]
    [InlineData("DEC2OCT(-100)", "7777777634")]
    [InlineData("BIN2HEX(11111011,4)", "00FB")]
    [InlineData("BIN2HEX(1111111111)", "FFFFFFFFFF")]
    [InlineData("BIN2OCT(1001,3)", "011")]
    [InlineData("BIN2OCT(1111111111)", "7777777777")]
    [InlineData("HEX2BIN(\"F\",8)", "00001111")]
    [InlineData("HEX2BIN(\"B7\")", "10110111")]
    [InlineData("HEX2BIN(\"FFFFFFFFFF\")", "1111111111")]
    [InlineData("HEX2OCT(\"F\",3)", "017")]
    [InlineData("HEX2OCT(\"3B4E\")", "35516")]
    [InlineData("HEX2OCT(\"FFFFFFFF00\")", "7777777400")]
    [InlineData("OCT2BIN(3,3)", "011")]
    [InlineData("OCT2BIN(7777777000)", "1000000000")]
    [InlineData("OCT2HEX(100,4)", "0040")]
    [InlineData("OCT2HEX(7777777533)", "FFFFFFFF5B")]
    public void BasesFromDecimal(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Fact]
    public void BasesRefuseWhatDoesNotFit()
    {
        Assert.Equal(ExcelError.Number, Eval("DEC2BIN(512)").GetError());
        Assert.Equal(ExcelError.Number, Eval("DEC2BIN(9,2)").GetError());
        Assert.Equal(ExcelError.Number, Eval("BIN2DEC(\"12\")").GetError());
        Assert.Equal(ExcelError.Number, Eval("HEX2DEC(\"FFFFFFFFFFF\")").GetError());
    }

    [Theory]
    [InlineData("DELTA(5,4)", 0.0)]
    [InlineData("DELTA(5,5)", 1.0)]
    [InlineData("DELTA(0.5,0)", 0.0)]
    [InlineData("GESTEP(5,4)", 1.0)]
    [InlineData("GESTEP(5,5)", 1.0)]
    [InlineData("GESTEP(-4,-5)", 1.0)]
    [InlineData("GESTEP(-1)", 0.0)]
    public void StepFunctions(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber());
    }

    [Theory]
    [InlineData("CONVERT(1,\"lbm\",\"kg\")", 0.45359237)]
    [InlineData("CONVERT(68,\"F\",\"C\")", 20.0)]
    [InlineData("CONVERT(100,\"C\",\"F\")", 212.0)]
    [InlineData("CONVERT(2.5,\"ft\",\"in\")", 30.0)]
    [InlineData("CONVERT(1,\"gal\",\"l\")", 3.785411784)]
    [InlineData("CONVERT(1,\"day\",\"hr\")", 24.0)]
    [InlineData("CONVERT(1,\"Mibyte\",\"bit\")", 8388608.0)]
    [InlineData("CONVERT(60,\"mph\",\"m/s\")", 26.8224)]
    public void ConvertBetweenUnits(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void ConvertRefusesCrossedCategoriesAndStrangers()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("CONVERT(2.5,\"ft\",\"sec\")").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("CONVERT(1,\"nonsense\",\"m\")").GetError());

        // Unit names are case-sensitive: a lowercase kelvin is no unit at all.
        Assert.Equal(ExcelError.NotAvailable, Eval("CONVERT(1,\"k\",\"C\")").GetError());
    }

    [Theory]
    [InlineData("COMPLEX(3,4)", "3+4i")]
    [InlineData("COMPLEX(3,-4,\"j\")", "3-4j")]
    [InlineData("COMPLEX(0,1)", "i")]
    [InlineData("COMPLEX(1,0)", "1")]
    [InlineData("COMPLEX(0,0)", "0")]
    [InlineData("IMCONJUGATE(\"3+4i\")", "3-4i")]
    [InlineData("IMSUM(\"3+4i\",\"5-3i\")", "8+i")]
    [InlineData("IMSUB(\"13+4i\",\"5+3i\")", "8+i")]
    [InlineData("IMPRODUCT(\"3+4i\",\"5-3i\")", "27+11i")]
    [InlineData("IMDIV(\"-238+240i\",\"10+24i\")", "5+12i")]
    public void ComplexArithmetic(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Theory]
    [InlineData("IMREAL(\"6-9i\")", 6.0)]
    [InlineData("IMAGINARY(\"3+4i\")", 4.0)]
    [InlineData("IMAGINARY(\"0-j\")", -1.0)]
    [InlineData("IMAGINARY(4)", 0.0)]
    [InlineData("IMABS(\"5+12i\")", 13.0)]
    [InlineData("IMARGUMENT(\"3+4i\")", 0.92729522)]
    [InlineData("IMREAL(IMPOWER(\"2+3i\",3))", -46.0)]
    [InlineData("IMAGINARY(IMPOWER(\"2+3i\",3))", 9.0)]
    [InlineData("IMREAL(IMSQRT(\"1+i\"))", 1.09868411)]
    [InlineData("IMAGINARY(IMSQRT(\"1+i\"))", 0.45508986)]
    [InlineData("IMREAL(IMEXP(\"1+i\"))", 1.46869394)]
    [InlineData("IMAGINARY(IMLN(\"3+4i\"))", 0.92729522)]
    [InlineData("IMREAL(IMSIN(\"3+4i\"))", 3.85373804)]
    public void ComplexParts(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Fact]
    public void ComplexRefusesMixedSuffixes()
    {
        Assert.Equal(ExcelError.Value, Eval("IMSUM(\"3+4i\",\"5-3j\")").GetError());
        Assert.Equal(ExcelError.Value, Eval("COMPLEX(1,2,\"k\")").GetError());
        Assert.Equal(ExcelError.DivisionByZero, Eval("IMARGUMENT(\"0\")").GetError());
    }

    [Theory]
    [InlineData("BESSELJ(1.9,2)", 0.32992576)]
    [InlineData("BESSELI(1.5,1)", 0.98166643)]
    [InlineData("BESSELK(1.5,1)", 0.27738780)]
    [InlineData("BESSELY(2.5,1)", 0.14591814)]
    [InlineData("BESSELJ(10,0)", -0.24593576)]
    [InlineData("BESSELI(10,0)", 2815.71662847)]
    [InlineData("BESSELY(10,0)", 0.05567117)]
    // J₂(25) checked through the recurrence J₂ = (2/x)·J₁ − J₀ on the tabulated J₀, J₁.
    [InlineData("BESSELJ(25,2)", -0.10629480)]
    public void BesselFamilies(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    [Fact]
    public void BesselDomains()
    {
        Assert.Equal(ExcelError.Number, Eval("BESSELJ(1,-1)").GetError());
        Assert.Equal(ExcelError.Number, Eval("BESSELY(-1,0)").GetError());
        Assert.Equal(ExcelError.Number, Eval("BESSELK(-1,0)").GetError());

        // J and I take the negative axis by symmetry.
        Assert.Equal(Eval("BESSELJ(2,0)").GetNumber(), Eval("BESSELJ(-2,0)").GetNumber(), precision: 12);
        Assert.Equal(-Eval("BESSELJ(2,1)").GetNumber(), Eval("BESSELJ(-2,1)").GetNumber(), precision: 12);
    }

    [Fact]
    public void BesselKMatchesTheReferenceValue()
    {
        Assert.Equal(1.778006232e-5, Eval("BESSELK(10,0)").GetNumber(), precision: 12);
    }
}
