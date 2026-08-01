using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The statistical distributions, against the values Excel's own documentation publishes for
/// the same arguments. The inverses are checked by round trip as well, since an inverse that
/// agrees with its own forward function to fifteen digits cannot be far wrong.
/// </summary>
public sealed class DistributionTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        return sheet.Evaluate(formula, "A1");
    }

    [Theory]
    [InlineData("NORM.DIST(42,40,1.5,TRUE)", 0.9087888)]
    [InlineData("NORM.DIST(42,40,1.5,FALSE)", 0.10934005)]
    [InlineData("NORM.S.DIST(1.333333,TRUE)", 0.90878870)]
    [InlineData("NORM.S.DIST(1.333333,FALSE)", 0.16401011)]
    [InlineData("NORMSDIST(1.333333)", 0.90878870)]
    [InlineData("STANDARDIZE(42,40,1.5)", 1.33333333)]
    [InlineData("PHI(0.75)", 0.30113743)]
    [InlineData("GAUSS(2)", 0.47724987)]
    [InlineData("LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.03908363)]
    [InlineData("LOGNORMDIST(4,3.5,1.2)", 0.03908363)]
    [InlineData("EXPON.DIST(0.2,10,TRUE)", 0.86466472)]
    [InlineData("EXPON.DIST(0.2,10,FALSE)", 1.35335283)]
    [InlineData("WEIBULL.DIST(105,20,100,TRUE)", 0.92958139)]
    [InlineData("GAMMA(2.5)", 1.32934039)]
    [InlineData("GAMMALN(4)", 1.79175947)]
    [InlineData("ERF(0.745)", 0.70792892)]
    [InlineData("ERF(1)", 0.84270079)]
    [InlineData("ERF(1,2)", 0.15262147)]
    [InlineData("ERFC(1)", 0.15729921)]
    [InlineData("ERF.PRECISE(1)", 0.84270079)]
    [InlineData("ERFC.PRECISE(1)", 0.15729921)]
    public void TheContinuousFamiliesAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("CHISQ.DIST(0.5,1,TRUE)", 0.52049988)]
    [InlineData("CHISQ.DIST(0.5,1,FALSE)", 0.43939129)]
    [InlineData("T.DIST(60,1,TRUE)", 0.99469533)]
    [InlineData("T.DIST.2T(1.959999998,60)", 0.05464493)]
    [InlineData("T.DIST.RT(1.959999998,60)", 0.02732246)]
    [InlineData("TDIST(1.959999998,60,2)", 0.05464493)]
    [InlineData("T.INV(0.75,2)", 0.81649658)]
    [InlineData("F.DIST.RT(15.20686486,6,4)", 0.01)]
    [InlineData("FDIST(15.20686486,6,4)", 0.01)]
    [InlineData("F.DIST(15.20686486,6,4,TRUE)", 0.99)]
    public void TheTestStatisticsAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("BINOM.DIST(6,10,0.5,FALSE)", 0.20507813)]
    [InlineData("BINOM.DIST(6,10,0.5,TRUE)", 0.828125)]
    [InlineData("BINOMDIST(6,10,0.5,TRUE)", 0.828125)]
    [InlineData("BINOM.INV(6,0.5,0.75)", 4.0)]
    [InlineData("CRITBINOM(6,0.5,0.75)", 4.0)]
    [InlineData("NEGBINOM.DIST(10,5,0.25,FALSE)", 0.05504866)]
    [InlineData("NEGBINOMDIST(10,5,0.25)", 0.05504866)]
    [InlineData("POISSON.DIST(2,5,TRUE)", 0.12465202)]
    [InlineData("POISSON.DIST(2,5,FALSE)", 0.08422434)]
    [InlineData("POISSON(2,5,FALSE)", 0.08422434)]
    [InlineData("HYPGEOM.DIST(1,4,8,20,FALSE)", 0.36326109)]
    [InlineData("HYPGEOM.DIST(1,4,8,20,TRUE)", 0.46542827)]
    [InlineData("HYPGEOMDIST(1,4,8,20)", 0.36326109)]
    public void TheCountingFamiliesAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    [Theory]
    [InlineData("BETA.DIST(2,8,10,TRUE,1,3)", 0.68547058)]
    [InlineData("BETA.DIST(2,8,10,FALSE,1,3)", 1.48376464)]
    [InlineData("BETADIST(2,8,10,1,3)", 0.68547058)]
    [InlineData("GAMMA.DIST(10.00001131,9,2,FALSE)", 0.03263913)]
    public void TheTwoShapedFamiliesAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    /// <summary>
    /// The cases Excel's documentation only prints to six decimals. Asserting more than it
    /// publishes would be asserting digits nobody has checked; the accuracy underneath is
    /// pinned by <see cref="TheIncompleteGammaMatchesItsClosedForm"/> instead.
    /// </summary>
    [Theory]
    [InlineData("CHISQ.DIST.RT(18.307,10)", 0.050001)]
    [InlineData("CHIDIST(18.307,10)", 0.050001)]
    [InlineData("GAMMA.DIST(10.00001131,9,2,TRUE)", 0.068094)]
    [InlineData("GAMMADIST(10.00001131,9,2,TRUE)", 0.068094)]
    public void TheSixDecimalCasesAgreeWithExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// Three exact points that need no published figure at all: a symmetric distribution is
    /// half spent at its own centre, and an F with matching degrees of freedom has its median
    /// at one.
    /// </summary>
    [Theory]
    [InlineData("NORM.S.DIST(0,TRUE)", 0.5)]
    [InlineData("T.DIST(0,7,TRUE)", 0.5)]
    [InlineData("F.DIST(1,5,5,TRUE)", 0.5)]
    [InlineData("BETA.DIST(0.5,3,3,TRUE)", 0.5)]
    [InlineData("ERF(0)", 0.0)]
    [InlineData("ERFC(0)", 1.0)]
    public void TheCentreOfEachDistributionIsExact(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 14);
    }

    /// <summary>
    /// At a whole number of degrees of freedom the chi-squared tail has a closed form — a
    /// finite sum, no series and no continued fraction — so the two must agree to the last
    /// digit. Nothing else here checks the incomplete gamma that closely.
    /// </summary>
    [Theory]
    [InlineData(9.1535)]
    [InlineData(30.0)]
    [InlineData(0.5)]
    [InlineData(4.0)]
    public void TheIncompleteGammaMatchesItsClosedForm(double half)
    {
        // Ten degrees of freedom means a shape of five, and Q(5,x) = e^-x (1+x+x²/2+x³/6+x⁴/24).
        string closed = $"EXP(-{half})*(1+{half}+{half}^2/2+{half}^3/6+{half}^4/24)";
        double expected = Eval(closed).GetNumber();
        double actual = Eval($"CHISQ.DIST.RT({2 * half},10)").GetNumber();

        Assert.True(
            Math.Abs(actual - expected) <= expected * 1e-14,
            $"tail at {half} gave {actual:E17}, closed form {expected:E17}");
    }

    [Theory]
    [InlineData("CONFIDENCE(0.05,2.5,50)", 0.69295191)]
    [InlineData("CONFIDENCE.NORM(0.05,2.5,50)", 0.69295191)]
    [InlineData("CONFIDENCE.T(0.05,1,50)", 0.284196855)]
    public void ConfidenceWidensWhenTheSpreadWasEstimated(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 7);
    }

    /// <summary>
    /// Every inverse undoes its own forward function. That catches a mistake in either half
    /// without depending on the last digit of a published figure.
    /// </summary>
    [Theory]
    [InlineData("NORM.INV(NORM.DIST(42,40,1.5,TRUE),40,1.5)", 42.0)]
    [InlineData("NORM.S.INV(NORM.S.DIST(1.5,TRUE))", 1.5)]
    [InlineData("NORMSINV(NORMSDIST(-0.8))", -0.8)]
    [InlineData("LOGNORM.INV(LOGNORM.DIST(4,3.5,1.2,TRUE),3.5,1.2)", 4.0)]
    [InlineData("LOGINV(LOGNORMDIST(4,3.5,1.2),3.5,1.2)", 4.0)]
    [InlineData("GAMMA.INV(GAMMA.DIST(7,9,2,TRUE),9,2)", 7.0)]
    [InlineData("BETA.INV(BETA.DIST(2,8,10,TRUE,1,3),8,10,1,3)", 2.0)]
    [InlineData("BETAINV(BETADIST(0.4,2,3),2,3)", 0.4)]
    [InlineData("CHISQ.INV(CHISQ.DIST(3,4,TRUE),4)", 3.0)]
    [InlineData("CHISQ.INV.RT(CHISQ.DIST.RT(18.307,10),10)", 18.307)]
    [InlineData("CHIINV(CHIDIST(18.307,10),10)", 18.307)]
    [InlineData("T.INV(T.DIST(1.4,7,TRUE),7)", 1.4)]
    [InlineData("T.INV.2T(T.DIST.2T(1.96,60),60)", 1.96)]
    [InlineData("TINV(TDIST(1.96,60,2),60)", 1.96)]
    [InlineData("F.INV(F.DIST(2.5,6,4,TRUE),6,4)", 2.5)]
    [InlineData("F.INV.RT(F.DIST.RT(15.20686486,6,4),6,4)", 15.20686486)]
    [InlineData("FINV(FDIST(15.20686486,6,4),6,4)", 15.20686486)]
    public void EveryInverseUndoesItsOwnDistribution(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 6);
    }

    /// <summary>
    /// The far tail is where an implementation that subtracts one distribution from a whole
    /// falls apart: every digit of the answer is lost to rounding. These come from the tail's
    /// own expansion, so they keep their precision all the way down.
    /// </summary>
    [Theory]
    [InlineData("NORM.S.DIST(-5,TRUE)", 2.8665157187919391e-7)]
    [InlineData("NORM.S.DIST(-6,TRUE)", 9.8658764503769814e-10)]
    [InlineData("NORM.S.DIST(-8,TRUE)", 6.2209605742716393e-16)]
    [InlineData("CHISQ.DIST.RT(60,10)", 3.6243009520614915e-9)]
    public void TheFarTailKeepsItsDigits(string formula, double expected)
    {
        double actual = Eval(formula).GetNumber();
        Assert.True(
            Math.Abs(actual - expected) <= Math.Abs(expected) * 1e-10,
            $"{formula} gave {actual:E17}, expected {expected:E17}");
    }

    /// <summary>Arguments that leave a distribution undefined are refused rather than guessed at.</summary>
    [Theory]
    [InlineData("NORM.DIST(1,0,0,TRUE)")]
    [InlineData("NORM.DIST(1,0,-1,TRUE)")]
    [InlineData("NORM.INV(0,0,1)")]
    [InlineData("NORM.INV(1,0,1)")]
    [InlineData("LOGNORM.DIST(0,0,1,TRUE)")]
    [InlineData("EXPON.DIST(-1,10,TRUE)")]
    [InlineData("EXPON.DIST(1,0,TRUE)")]
    [InlineData("WEIBULL.DIST(1,0,100,TRUE)")]
    [InlineData("GAMMA.DIST(1,0,2,TRUE)")]
    [InlineData("BETA.DIST(4,8,10,TRUE,1,3)")]
    [InlineData("CHISQ.DIST(-1,10,TRUE)")]
    [InlineData("CHISQ.DIST(1,0,TRUE)")]
    [InlineData("T.DIST(1,0,TRUE)")]
    [InlineData("T.DIST.2T(-1,10)")]
    [InlineData("TDIST(-1,10,2)")]
    [InlineData("TDIST(1,10,3)")]
    [InlineData("F.DIST(1,0,4,TRUE)")]
    [InlineData("BINOM.DIST(11,10,0.5,TRUE)")]
    [InlineData("BINOM.DIST(1,10,1.5,TRUE)")]
    [InlineData("POISSON.DIST(-1,5,TRUE)")]
    [InlineData("HYPGEOM.DIST(5,4,8,20,FALSE)")]
    [InlineData("CONFIDENCE(0,2.5,50)")]
    [InlineData("CONFIDENCE.T(0.05,1,1)")]
    public void ImpossibleArgumentsAreRefused(string formula)
    {
        Assert.Equal(ExcelError.Number, Eval(formula).GetError());
    }
}
