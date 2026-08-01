using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The places where Excel's functions differ from what ISO/IEC 29500 says they do, taken from
/// the deviations Microsoft documents in MS-OI29500 section 2.1.
/// </summary>
/// <remarks>
/// These matter because the standard is the wrong reference for most of them. A file is written
/// by Excel and read by a user who expects Excel's answer, so where the two disagree the
/// standard's reading produces a defensible result that is nonetheless not the one in the cell.
/// Every case here names the deviation it pins.
/// </remarks>
public sealed class ConformanceTests
{
    private static Worksheet Sheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet["A2"] = 2;
        sheet["A3"] = 3;
        sheet["A4"] = 4;
        sheet["B1"] = "text";
        return sheet;
    }

    private static CellValue Eval(string formula) => Sheet().Evaluate(formula, "Z1");

    private static ExcelError ErrorOf(string formula)
    {
        CellValue value = Eval(formula);
        Assert.True(value.IsError, $"{formula} returned {value} rather than an error.");
        return value.GetError();
    }

    /// <summary>
    /// Errors Excel returns where the standard names a different one, or names none at all.
    /// </summary>
    [Theory]
    // 2.1.877 ADDRESS: the standard says #NUM! for a row or column below one.
    [InlineData("ADDRESS(0,1)", ExcelError.Value)]
    [InlineData("ADDRESS(1,0)", ExcelError.Value)]
    // 2.1.895 CEILING and 2.1.954 FLOOR: only a positive number with a negative step refuses.
    [InlineData("CEILING(2.5,-2)", ExcelError.Number)]
    [InlineData("FLOOR(2.5,-2)", ExcelError.Number)]
    // 2.1.986 LOG: the standard does not say what a base of one does.
    [InlineData("LOG(10,1)", ExcelError.DivisionByZero)]
    // 2.1.1018 POWER: the standard says #DIV/0! for zero to the zero.
    [InlineData("POWER(0,0)", ExcelError.Number)]
    // 2.1.990 MATCH: the standard says #NUM! when nothing matches.
    [InlineData("MATCH(99,A1:A4,0)", ExcelError.NotAvailable)]
    // 2.1.1049 SUMPRODUCT: the standard says #NUM! for mismatched shapes.
    [InlineData("SUMPRODUCT(A1:A4,A1:A2)", ExcelError.Value)]
    // 2.1.1046 SUBSTITUTE: the standard admits an occurrence of zero.
    [InlineData("SUBSTITUTE(\"aaa\",\"a\",\"b\",0)", ExcelError.Value)]
    // 2.1.992 MID and 2.1.950 FIND: the standard counts positions from zero.
    [InlineData("MID(\"abc\",0,2)", ExcelError.Value)]
    [InlineData("FIND(\"a\",\"abc\",0)", ExcelError.Value)]
    [InlineData("SEARCH(\"a\",\"abc\",0)", ExcelError.Value)]
    [InlineData("REPLACE(\"abc\",0,1,\"x\")", ExcelError.Value)]
    // 2.1.1042 STDEV and 2.1.1062 VAR: the standard does not say what one sample does.
    [InlineData("STDEV(A1)", ExcelError.DivisionByZero)]
    [InlineData("VAR(A1)", ExcelError.DivisionByZero)]
    // 2.1.887 AVERAGEIF: the standard does not say what matching nothing does.
    [InlineData("AVERAGEIF(A1:A4,\">100\")", ExcelError.DivisionByZero)]
    // 2.1.1016 PMT: the standard places no restriction on a term of zero periods.
    [InlineData("PMT(0.05,0,1000)", ExcelError.DivisionByZero)]
    // 2.1.1071 XNPV: the standard places no restriction on the rate.
    [InlineData("XNPV(-0.1,{-100;120},{1;40})", ExcelError.Number)]
    // 2.1.1026 RANK: the standard does not say what ranking an absent value does.
    [InlineData("RANK(99,A1:A4)", ExcelError.NotAvailable)]
    // 2.1.1003 NORMSINV and 2.1.999 NEGBINOMDIST: the standard admits the endpoints.
    [InlineData("NORMSINV(0)", ExcelError.Number)]
    [InlineData("NORMSINV(1)", ExcelError.Number)]
    [InlineData("LOGINV(0,0,1)", ExcelError.Number)]
    public void OfficeReturnsTheErrorItDocuments(string formula, ExcelError expected)
    {
        Assert.Equal(expected, ErrorOf(formula));
    }

    /// <summary>
    /// 2.1.898 CLEAN. The standard strips only the C0 control characters; Excel strips the C1
    /// range as well, which is where a document pasted out of a word processor keeps its
    /// curly-quote artefacts.
    /// </summary>
    [Fact]
    public void CleanStripsBothControlRanges()
    {
        Assert.Equal("ab", Eval("CLEAN(\"a\" & CHAR(7) & \"b\")").GetText());
        Assert.Equal("ab", Eval("CLEAN(\"a\" & CHAR(133) & \"b\")").GetText());
    }

    /// <summary>
    /// 2.1.962 HOUR. The standard gives the range as 0 to 59, which is the range of a minute;
    /// Excel returns an hour.
    /// </summary>
    [Fact]
    public void HourIsAnHourAndNotAMinute()
    {
        Assert.Equal(23, Eval("HOUR(0.9999)").GetNumber());
    }

    /// <summary>
    /// 2.1.1054 TIME and 2.1.1055 TIMEVALUE. A whole day is not a time: the serial value is
    /// below one, and a time of exactly 24 hours wraps to zero.
    /// </summary>
    [Fact]
    public void ATimeStaysBelowAWholeDay()
    {
        Assert.Equal(0, Eval("TIME(24,0,0)").ToSerialNumber());
        Assert.True(Eval("TIMEVALUE(\"23:59:59\")").ToSerialNumber() < 1);
    }

    /// <summary>
    /// 2.1.946 EDATE. Adding a month to the 31st lands on a month that has no 31st, and Excel
    /// takes the last day of it rather than spilling into the next.
    /// </summary>
    [Fact]
    public void EdateStopsAtTheEndOfTheMonth()
    {
        Assert.Equal(new DateTime(2026, 2, 28), Eval("EDATE(DATE(2026,1,31),1)").GetDateTime());
        Assert.Equal(new DateTime(2024, 2, 29), Eval("EDATE(DATE(2024,1,31),1)").GetDateTime());
    }

    /// <summary>
    /// 2.1.1000 NETWORKDAYS. Both ends count, and the time of day is ignored — a Monday to the
    /// same Monday is one working day, not none.
    /// </summary>
    [Fact]
    public void WorkingDaysCountBothEnds()
    {
        Assert.Equal(1, Eval("NETWORKDAYS(DATE(2026,7,27),DATE(2026,7,27))").GetNumber());
        Assert.Equal(5, Eval("NETWORKDAYS(DATE(2026,7,27),DATE(2026,7,31))").GetNumber());
    }

    /// <summary>
    /// 2.1.923 DAYS360. The American method moves the end of February to the thirtieth as well,
    /// which the standard describes for the 31st only.
    /// </summary>
    [Fact]
    public void TheAmericanMonthEndsOnTheThirtieth()
    {
        Assert.Equal(30, Eval("DAYS360(DATE(2026,2,28),DATE(2026,3,30))").GetNumber());
    }

    /// <summary>
    /// 2.1.1017 POISSON. The standard refuses a mean of zero; Excel allows it, and the
    /// distribution it describes is a real one — every outcome is zero events.
    /// </summary>
    [Fact]
    public void APoissonMeanOfZeroIsAllowed()
    {
        Assert.Equal(1, Eval("POISSON(0,0,TRUE)").GetNumber());
        Assert.Equal(ExcelError.Number, ErrorOf("POISSON(0,-1,TRUE)"));
    }

    /// <summary>
    /// 2.1.1014 PERCENTILE. The standard leaves an empty array undefined, which is the one
    /// answer a spreadsheet cannot give.
    /// </summary>
    [Fact]
    public void APercentileOfNothingIsZero()
    {
        Assert.Equal(0, Eval("PERCENTILE(D1:D4,0.5)").GetNumber());
    }

    /// <summary>
    /// 2.1.974 INDIRECT. A reference longer than 255 characters is refused rather than parsed,
    /// which also caps the work a hostile formula can ask for.
    /// </summary>
    [Fact]
    public void AnOverlongIndirectReferenceIsRefused()
    {
        Assert.Equal(ExcelError.Value, ErrorOf($"INDIRECT(\"{new string('a', 256)}\")"));
    }
}
