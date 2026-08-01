using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The Excel behaviours that separate a working engine from an almost-working one: blank
/// handling, the ordering of types in comparisons, rounding direction and error propagation.
/// </summary>
public sealed class SemanticsTests
{
    private static CellValue Eval(string formula)
    {
        var book = Workbook.Create();
        book.Sheets.Add("S");
        return book.Evaluate(formula);
    }

    private static CellValue EvalOnSheet(Action<Worksheet> arrange, string formula, string origin = "Z1")
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        arrange(sheet);
        return sheet.Evaluate(formula, origin);
    }

    [Theory]
    [InlineData("1+1", 2.0)]
    [InlineData("2*3", 6.0)]
    [InlineData("7/2", 3.5)]
    [InlineData("2^10", 1024.0)]
    [InlineData("-2^2", 4.0)]
    [InlineData("10%", 0.1)]
    [InlineData("1+2*3", 7.0)]
    [InlineData("(1+2)*3", 9.0)]
    [InlineData("2^3^2", 64.0)]
    public void Arithmetic_FollowsExcelPrecedence(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 10);
    }

    [Theory]
    [InlineData("\"1\"+1", 2.0)]
    [InlineData("TRUE+1", 2.0)]
    [InlineData("\"1.5\"*2", 3.0)]
    [InlineData("\"50%\"+0", 0.5)]
    public void TextAndLogicals_CoerceInArithmetic(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 10);
    }

    [Theory]
    [InlineData("\"a\"&\"b\"", "ab")]
    [InlineData("1&2", "12")]
    [InlineData("TRUE&\"\"", "TRUE")]
    [InlineData("1.5&\"\"", "1.5")]
    public void Concatenation_UsesGeneralFormatting(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    [Fact]
    public void BlankCells_ActAsZeroOrEmptyText()
    {
        Assert.Equal(1.0, EvalOnSheet(static _ => { }, "A1+1").GetNumber());
        Assert.True(EvalOnSheet(static _ => { }, "A1=0").GetBoolean());
        Assert.True(EvalOnSheet(static _ => { }, "A1=\"\"").GetBoolean());
        Assert.True(EvalOnSheet(static _ => { }, "ISBLANK(A1)").GetBoolean());
        Assert.False(EvalOnSheet(static s => s["A1"] = "", "ISBLANK(A1)").GetBoolean());
    }

    [Theory]
    [InlineData("1<\"a\"", true)]        // Numbers sort before text.
    [InlineData("\"a\"<TRUE", true)]     // Text sorts before logicals.
    [InlineData("\"a\"=\"A\"", true)]    // Text comparison ignores case.
    [InlineData("EXACT(\"a\",\"A\")", false)]
    [InlineData("FALSE<TRUE", true)]
    public void Comparisons_UseExcelTypeOrdering(string formula, bool expected)
    {
        Assert.Equal(expected, Eval(formula).GetBoolean());
    }

    [Theory]
    [InlineData("ROUND(2.5,0)", 3.0)]      // Away from zero, not banker's rounding.
    [InlineData("ROUND(-2.5,0)", -3.0)]
    [InlineData("ROUND(1.005,2)", 1.01)]
    [InlineData("ROUNDUP(1.001,2)", 1.01)]
    [InlineData("ROUNDDOWN(1.999,2)", 1.99)]
    [InlineData("TRUNC(-1.9)", -1.0)]
    [InlineData("INT(-1.9)", -2.0)]
    [InlineData("MOD(-3,2)", 1.0)]        // MOD takes the divisor's sign.
    [InlineData("MOD(3,-2)", -1.0)]
    public void Rounding_MatchesExcel(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 10);
    }

    [Theory]
    [InlineData("1/0", ExcelError.DivisionByZero)]
    [InlineData("MOD(1,0)", ExcelError.DivisionByZero)]
    [InlineData("\"x\"+1", ExcelError.Value)]
    [InlineData("SQRT(-1)", ExcelError.Number)]
    [InlineData("LOG(0)", ExcelError.Number)]
    [InlineData("NOSUCHFUNC(1)", ExcelError.Name)]
    [InlineData("UnknownName", ExcelError.Name)]
    [InlineData("NA()", ExcelError.NotAvailable)]
    public void Errors_UseTheRightCode(string formula, ExcelError expected)
    {
        Assert.Equal(expected, Eval(formula).GetError());
    }

    [Fact]
    public void Errors_PropagateThroughOperatorsButNotThroughIfError()
    {
        Assert.Equal(ExcelError.DivisionByZero, Eval("1/0+1").GetError());
        Assert.Equal("caught", Eval("IFERROR(1/0,\"caught\")").GetText());
        Assert.Equal(5.0, Eval("IFERROR(5,\"caught\")").GetNumber());
        Assert.True(Eval("ISERROR(1/0)").GetBoolean());
        Assert.False(Eval("ISERR(NA())").GetBoolean());
        Assert.True(Eval("ISNA(NA())").GetBoolean());
    }

    [Fact]
    public void IfChoosesOneBranch_AndDiscardsErrorsFromTheOther()
    {
        Assert.Equal(0.0, EvalOnSheet(static s => s["A1"] = 0, "IF(A1=0,0,1/A1)").GetNumber());
    }

    [Fact]
    public void TextResultsBeyondTheCellLimit_BecomeValueError()
    {
        Assert.Equal(ExcelError.Value, Eval("REPT(\"ab\",20000)").GetError());
    }

    [Fact]
    public void WholeColumnReferences_ClampToTheUsedRange()
    {
        CellValue sum = EvalOnSheet(
            static sheet =>
            {
                sheet["A1"] = 1;
                sheet["A2"] = 2;
                sheet["A3"] = 3;
            },
            "SUM(A:A)");

        Assert.Equal(6.0, sum.GetNumber());
    }

    [Fact]
    public void ImplicitIntersection_PicksTheFormulaRowOrColumn()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 10;
        sheet["A2"] = 20;
        sheet["A3"] = 30;

        // A formula sitting in row 2 sees A2 when it references the whole column range.
        Assert.Equal(20.0, sheet.Evaluate("A1:A3", "B2").GetNumber());
        Assert.Equal(ExcelError.Value, sheet.Evaluate("A1:A3", "B9").GetError());
    }

    [Fact]
    public void AggregatesSkipTextInRangesButConvertLiterals()
    {
        Assert.Equal(1.0, Eval("SUM(TRUE)").GetNumber());
        Assert.Equal(3.0, Eval("SUM(\"1\",2)").GetNumber());

        CellValue fromRange = EvalOnSheet(
            static sheet =>
            {
                sheet["A1"] = 1;
                sheet["A2"] = "text";
                sheet["A3"] = true;
            },
            "SUM(A1:A3)");
        Assert.Equal(1.0, fromRange.GetNumber());
    }

    [Fact]
    public void CountFamily_DistinguishesBlanksFromEmptyText()
    {
        static void Arrange(Worksheet sheet)
        {
            sheet["A1"] = 1;
            sheet["A2"] = "";
            sheet["A4"] = "text";
        }

        Assert.Equal(1.0, EvalOnSheet(Arrange, "COUNT(A1:A4)").GetNumber());
        Assert.Equal(3.0, EvalOnSheet(Arrange, "COUNTA(A1:A4)").GetNumber());
        Assert.Equal(2.0, EvalOnSheet(Arrange, "COUNTBLANK(A1:A4)").GetNumber());
    }

    [Fact]
    public void DatesRoundTripThroughSerialArithmetic()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = new DateTime(2026, 7, 24);

        Assert.Equal(2026.0, sheet.Evaluate("YEAR(A1)").GetNumber());
        Assert.Equal(7.0, sheet.Evaluate("MONTH(A1)").GetNumber());
        Assert.Equal(24.0, sheet.Evaluate("DAY(A1)").GetNumber());
        Assert.Equal(new DateTime(2026, 7, 31), sheet.Evaluate("EOMONTH(A1,0)").GetDateTime());
        Assert.Equal(1.0, sheet.Evaluate("A1+1-A1").GetNumber());
    }

    [Fact]
    public void Date1904System_ShiftsSerialsConsistently()
    {
        var book = Workbook.Create();
        book.DateSystem = ExcelDateSystem.Date1904;
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = new DateTime(2020, 1, 1);

        Assert.Equal(2020.0, sheet.Evaluate("YEAR(A1)").GetNumber());
        Assert.Equal(42_369.0, sheet.Evaluate("A1*1").GetNumber());
    }
}
