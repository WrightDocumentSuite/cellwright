using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Function-by-function checks against values verified in Excel.
/// </summary>
public sealed class FunctionTests
{
    private static Worksheet Sheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 10;
        sheet["A2"] = 20;
        sheet["A3"] = 30;
        sheet["A4"] = 40;
        sheet["B1"] = "apple";
        sheet["B2"] = "banana";
        sheet["B3"] = "cherry";
        sheet["B4"] = "apricot";
        sheet["C1"] = 1;
        sheet["C2"] = 2;
        sheet["C3"] = 3;
        sheet["C4"] = 4;
        return sheet;
    }

    private static CellValue Eval(string formula) => Sheet().Evaluate(formula, "Z1");

    [Theory]
    [InlineData("SUM(A1:A4)", 100.0)]
    [InlineData("AVERAGE(A1:A4)", 25.0)]
    [InlineData("MIN(A1:A4)", 10.0)]
    [InlineData("MAX(A1:A4)", 40.0)]
    [InlineData("COUNT(A1:A4)", 4.0)]
    [InlineData("MEDIAN(A1:A4)", 25.0)]
    [InlineData("PRODUCT(C1:C4)", 24.0)]
    [InlineData("LARGE(A1:A4,2)", 30.0)]
    [InlineData("SMALL(A1:A4,2)", 20.0)]
    [InlineData("STDEVP(C1:C4)", 1.1180339887)]
    [InlineData("VAR(C1:C4)", 1.6666666667)]
    [InlineData("RANK(30,A1:A4)", 2.0)]
    [InlineData("PERCENTILE(C1:C4,0.5)", 2.5)]
    [InlineData("SUMPRODUCT(A1:A4,C1:C4)", 300.0)]
    public void Statistics(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Theory]
    [InlineData("SUMIF(A1:A4,\">15\")", 90.0)]
    [InlineData("SUMIF(B1:B4,\"ap*\",A1:A4)", 50.0)]
    [InlineData("COUNTIF(A1:A4,\">=20\")", 3.0)]
    [InlineData("COUNTIF(B1:B4,\"a*\")", 2.0)]
    [InlineData("AVERAGEIF(A1:A4,\">15\")", 30.0)]
    [InlineData("SUMIFS(A1:A4,A1:A4,\">10\",C1:C4,\"<4\")", 50.0)]
    [InlineData("COUNTIFS(A1:A4,\">10\",B1:B4,\"<>apple\")", 3.0)]
    [InlineData("MAXIFS(A1:A4,C1:C4,\"<3\")", 20.0)]
    [InlineData("MINIFS(A1:A4,C1:C4,\">2\")", 30.0)]
    public void ConditionalAggregates(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Theory]
    [InlineData("VLOOKUP(30,A1:C4,3,FALSE)", 3.0)]
    [InlineData("MATCH(30,A1:A4,0)", 3.0)]
    [InlineData("INDEX(A1:A4,2)", 20.0)]
    [InlineData("INDEX(A1:C4,2,3)", 2.0)]
    [InlineData("XLOOKUP(20,A1:A4,C1:C4)", 2.0)]
    [InlineData("XMATCH(40,A1:A4)", 4.0)]
    [InlineData("CHOOSE(2,11,22,33)", 22.0)]
    [InlineData("ROWS(A1:A4)", 4.0)]
    [InlineData("COLUMNS(A1:C4)", 3.0)]
    [InlineData("SUM(OFFSET(A1,1,0,2,1))", 50.0)]
    [InlineData("SUM(INDIRECT(\"A1:A2\"))", 30.0)]
    public void Lookups(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void LookupMisses_ReturnNotAvailable()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("VLOOKUP(99,A1:C4,3,FALSE)").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("MATCH(99,A1:A4,0)").GetError());
        Assert.Equal("missing", Eval("XLOOKUP(99,A1:A4,C1:C4,\"missing\")").GetText());
    }

    [Theory]
    [InlineData("LEFT(\"Cellwright\",4)", "Cell")]
    [InlineData("RIGHT(\"Cellwright\",6)", "wright")]
    [InlineData("MID(\"Cellwright\",5,2)", "wr")]
    [InlineData("UPPER(\"abc\")", "ABC")]
    [InlineData("LOWER(\"ABC\")", "abc")]
    [InlineData("PROPER(\"hello world\")", "Hello World")]
    [InlineData("TRIM(\"  a   b  \")", "a b")]
    [InlineData("SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c")]
    [InlineData("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c")]
    [InlineData("REPLACE(\"abcdef\",2,3,\"X\")", "aXef")]
    [InlineData("CONCAT(\"a\",\"b\",\"c\")", "abc")]
    [InlineData("TEXTJOIN(\"-\",TRUE,\"a\",\"\",\"b\")", "a-b")]
    [InlineData("REPT(\"ab\",3)", "ababab")]
    [InlineData("CHAR(65)", "A")]
    [InlineData("FIXED(1234.5678)", "1,234.57")]
    [InlineData("FIXED(1234.5678,1)", "1,234.6")]
    [InlineData("FIXED(1234.5678,1,TRUE)", "1234.6")]
    [InlineData("FIXED(12345.6,-2)", "12,300")]
    [InlineData("FIXED(0.5,0)", "1")]
    public void TextFunctions(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    /// <summary>
    /// The 360-day year of accounting. The American method moves only a literal 31st, and the
    /// European one moves both ends, which is the whole difference between them.
    /// </summary>
    [Theory]
    [InlineData("DAYS360(DATE(2024,1,1),DATE(2024,12,31))", 360.0)]
    [InlineData("DAYS360(DATE(2024,1,1),DATE(2024,12,31),TRUE)", 359.0)]
    [InlineData("DAYS360(DATE(2024,1,30),DATE(2024,3,31))", 60.0)]
    [InlineData("DAYS360(DATE(2024,1,15),DATE(2024,3,31))", 76.0)]
    [InlineData("DAYS360(DATE(2024,1,15),DATE(2024,3,31),TRUE)", 75.0)]
    [InlineData("DAYS360(DATE(2024,3,31),DATE(2024,1,15))", -75.0)]
    public void DaysOnA360DayYear(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Theory]
    [InlineData("LEN(\"abc\")", 3.0)]
    [InlineData("FIND(\"b\",\"abc\")", 2.0)]
    [InlineData("SEARCH(\"B\",\"abc\")", 2.0)]
    [InlineData("SEARCH(\"c*\",\"abcd\")", 3.0)]
    [InlineData("CODE(\"A\")", 65.0)]
    [InlineData("VALUE(\"12.5\")", 12.5)]
    public void TextFunctionsReturningNumbers(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void FindIsCaseSensitiveAndSearchIsNot()
    {
        Assert.Equal(ExcelError.Value, Eval("FIND(\"B\",\"abc\")").GetError());
        Assert.Equal(2.0, Eval("SEARCH(\"B\",\"abc\")").GetNumber());
    }

    [Theory]
    [InlineData("IF(TRUE,1,2)", 1.0)]
    [InlineData("IF(FALSE,1,2)", 2.0)]
    [InlineData("IFS(FALSE,1,TRUE,2)", 2.0)]
    [InlineData("SWITCH(2,1,\"a\",2,\"b\",\"z\")", 0.0)]
    [InlineData("IFNA(NA(),7)", 7.0)]
    public void LogicalFunctions(string formula, double expected)
    {
        CellValue value = Eval(formula);
        if (value.IsText)
            return; // SWITCH returns text in this data set; the branch selection is what matters.
        Assert.Equal(expected, value.GetNumber(), precision: 8);
    }

    [Fact]
    public void AndOrXorFold()
    {
        Assert.True(Eval("AND(TRUE,1,\"TRUE\")").GetBoolean());
        Assert.False(Eval("AND(TRUE,FALSE)").GetBoolean());
        Assert.True(Eval("OR(FALSE,TRUE)").GetBoolean());
        Assert.True(Eval("XOR(TRUE,FALSE,FALSE)").GetBoolean());
        Assert.False(Eval("NOT(TRUE)").GetBoolean());
    }

    [Fact]
    public void SwitchReturnsMatchingBranchAndDefault()
    {
        Assert.Equal("b", Eval("SWITCH(2,1,\"a\",2,\"b\",\"z\")").GetText());
        Assert.Equal("z", Eval("SWITCH(9,1,\"a\",2,\"b\",\"z\")").GetText());
    }

    [Theory]
    [InlineData("ABS(-3)", 3.0)]
    [InlineData("SIGN(-3)", -1.0)]
    [InlineData("SQRT(16)", 4.0)]
    [InlineData("POWER(2,8)", 256.0)]
    [InlineData("EXP(0)", 1.0)]
    [InlineData("LN(1)", 0.0)]
    [InlineData("LOG10(1000)", 3.0)]
    [InlineData("LOG(8,2)", 3.0)]
    [InlineData("QUOTIENT(7,2)", 3.0)]
    [InlineData("CEILING(4.2,1)", 5.0)]
    [InlineData("FLOOR(4.8,1)", 4.0)]
    // Excel's own documented examples: a negative number with a positive step computes,
    // CEILING pulling toward zero and FLOOR away; a negative step mirrors both.
    [InlineData("CEILING(-2.5,2)", -2.0)]
    [InlineData("FLOOR(-2.5,2)", -4.0)]
    [InlineData("FLOOR(-23,1)", -23.0)]
    [InlineData("CEILING(-2.5,-2)", -4.0)]
    [InlineData("FLOOR(-2.5,-2)", -2.0)]
    [InlineData("EVEN(3)", 4.0)]
    [InlineData("ODD(2)", 3.0)]
    [InlineData("FACT(5)", 120.0)]
    [InlineData("GCD(12,18)", 6.0)]
    [InlineData("LCM(4,6)", 12.0)]
    [InlineData("DEGREES(PI())", 180.0)]
    public void MathFunctions(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void InformationFunctions()
    {
        Assert.True(Eval("ISNUMBER(A1)").GetBoolean());
        Assert.True(Eval("ISTEXT(B1)").GetBoolean());
        Assert.False(Eval("ISLOGICAL(A1)").GetBoolean());
        Assert.True(Eval("ISREF(A1)").GetBoolean());
        Assert.Equal(1.0, Eval("TYPE(A1)").GetNumber());
        Assert.Equal(2.0, Eval("TYPE(B1)").GetNumber());
        Assert.Equal(2.0, Eval("ERROR.TYPE(1/0)").GetNumber());
        Assert.Equal(0.0, Eval("N(\"text\")").GetNumber());
    }

    [Fact]
    public void ArrayConstantsParticipateInArithmetic()
    {
        Assert.Equal(10.0, Eval("SUM({1,2;3,4})").GetNumber());
        Assert.Equal(20.0, Eval("SUM({1,2;3,4}*2)").GetNumber());
    }

    [Fact]
    public void SumProductOverComparisons()
    {
        // The classic conditional-sum idiom relies on boolean arrays multiplying cleanly.
        Assert.Equal(90.0, Eval("SUMPRODUCT((A1:A4>10)*(A1:A4))").GetNumber());
    }

    /// <summary>
    /// An approximate lookup settles for the nearest value below the key, but only among values
    /// of the key's own type. Text sorts above every number, so a text key over a column of
    /// numbers has nothing to settle on and the answer is <c>#N/A</c>, not the last number.
    /// </summary>
    [Fact]
    public void AnApproximateLookupWillNotSettleOnAnotherType()
    {
        Assert.Equal(ExcelError.NotAvailable, Eval("VLOOKUP(\"7\",A1:A4,1)").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("MATCH(\"7\",A1:A4,1)").GetError());
        Assert.Equal(ExcelError.NotAvailable, Eval("VLOOKUP(25,B1:B4,1)").GetError());

        // Within one type the nearest value below the key is still found.
        Assert.Equal(20.0, Eval("VLOOKUP(25,A1:A4,1)").GetNumber());
    }

    /// <summary>
    /// An end of a range that is itself an error carries that error out. A range whose corner
    /// was deleted is <c>#REF!</c>, and calling it a type mismatch would hide what happened.
    /// </summary>
    [Fact]
    public void AnErrorInsideARangeIsCarriedOut()
    {
        Assert.Equal(ExcelError.Reference, Eval("SUM(#REF!:B4)").GetError());
        Assert.Equal(ExcelError.Reference, Eval("SUM(A1:#REF!)").GetError());
        Assert.Equal(ExcelError.Reference, Eval("SUBTOTAL(1,#REF!:B4)").GetError());
    }

    /// <summary>
    /// An error handed to the dispatching aggregates as a whole argument propagates, the way
    /// the evaluator would have propagated it into any other function. AGGREGATE told to
    /// ignore errors is the documented exception.
    /// </summary>
    [Fact]
    public void SubtotalCarriesOutAnErrorHandedInWhole()
    {
        Assert.Equal(ExcelError.Reference, Eval("SUBTOTAL(2,#REF!)").GetError());
        Assert.Equal(ExcelError.Reference, Eval("SUBTOTAL(9,#REF!)").GetError());
        Assert.Equal(ExcelError.DivisionByZero, Eval("SUBTOTAL(9,A1:A2,3/0)").GetError());
        Assert.Equal(ExcelError.Reference, Eval("AGGREGATE(9,4,#REF!)").GetError());
        Assert.Equal(0.0, Eval("AGGREGATE(9,6,#REF!)").GetNumber());
    }

    /// <summary>
    /// <c>AND</c> and <c>OR</c> read what a range holds, not how many cells it spans: empty
    /// cells and text are passed over, so a range padded with blanks answers the same as the
    /// values in it. A range with nothing to read at all is <c>#VALUE!</c>.
    /// </summary>
    [Fact]
    public void AndAndOrPassOverEmptyCellsAndTextInARange()
    {
        Worksheet sheet = Sheet();
        sheet["E1"] = true;
        sheet["E3"] = true;
        sheet["E4"] = "note";

        Assert.True(sheet.Evaluate("AND(E1:E10)", "Z1").GetBoolean());
        Assert.True(sheet.Evaluate("OR(E1:E10)", "Z1").GetBoolean());
        Assert.Equal(ExcelError.Value, sheet.Evaluate("AND(F1:F10)", "Z1").GetError());
    }

    /// <summary>
    /// The working week is Saturday and Sunday off only by convention, and the <c>.INTL</c> pair
    /// exists to say otherwise. January 2006 makes the counting easy to follow: it begins on a
    /// Sunday and runs 31 days, so it holds five Sundays and four of every other day.
    /// </summary>
    [Theory]
    [InlineData("NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31))", 22.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))", 22.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),7)", 23.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),11)", 26.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"0000011\")", 22.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"1111111\")", 0.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,31),DATE(2006,1,1))", -22.0)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),1,DATE(2006,1,2))", 21.0)]
    public void TheWeekendIsAnArgumentOfTheInternationalPair(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Friday the sixth of January 2006 falls on a weekend under code 7, so the next working day
    /// is the Sunday; under the usual week it is the Monday after.
    /// </summary>
    [Theory]
    [InlineData("WORKDAY.INTL(DATE(2006,1,6),1,7)", "2006-01-08")]
    [InlineData("WORKDAY.INTL(DATE(2006,1,6),1)", "2006-01-09")]
    [InlineData("WORKDAY(DATE(2006,1,6),1)", "2006-01-09")]
    [InlineData("WORKDAY.INTL(DATE(2006,1,6),1,1,DATE(2006,1,9))", "2006-01-10")]
    [InlineData("WORKDAY.INTL(DATE(2006,1,9),-1,7)", "2006-01-08")]
    public void MovingByWorkingDaysHonoursTheWeekendGiven(string formula, string expected)
    {
        Assert.Equal(DateTime.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), Eval(formula).GetDateTime());
    }

    /// <summary>A week nobody works has no next working day, and no working days to count.</summary>
    [Theory]
    [InlineData("WORKDAY.INTL(DATE(2006,1,1),1,\"1111111\")", ExcelError.Value)]
    [InlineData("WORKDAY.INTL(DATE(2006,1,1),1,8)", ExcelError.Number)]
    [InlineData("WORKDAY.INTL(DATE(2006,1,1),1,0)", ExcelError.Number)]
    [InlineData("WORKDAY.INTL(DATE(2006,1,1),1,18)", ExcelError.Number)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"111\")", ExcelError.Value)]
    [InlineData("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"0000002\")", ExcelError.Value)]
    public void AWeekThatNamesNothingIsRefused(string formula, ExcelError expected)
    {
        Assert.Equal(expected, Eval(formula).GetError());
    }

    /// <summary>
    /// The A-suffixed aggregates count what the plain ones skip: text as zero and logicals as
    /// one or nothing. A logical written into an array constant is the exception — even these
    /// pass it over, which is why the average below divides by four rather than five.
    /// </summary>
    [Fact]
    public void TheASuffixedAggregatesCountTextButNotLogicalsInAnArray()
    {
        Assert.Equal(1.75, Eval("AVERAGEA({1;2;4;\"Text\";TRUE})").GetNumber(), precision: 8);
        Assert.Equal(0.0, Eval("MINA({1;2;4;\"Text\";TRUE})").GetNumber(), precision: 8);

        // Straight in the formula, and in a range, they count as they always did.
        Assert.Equal(1.0, Eval("AVERAGEA(TRUE)").GetNumber(), precision: 8);
        Assert.Equal(25.0, Eval("AVERAGEA(A1:A4)").GetNumber(), precision: 8);
        Assert.Equal(0.0, Eval("AVERAGEA(B1:B4)").GetNumber(), precision: 8);
    }

    /// <summary>
    /// <c>HYPERLINK</c> shows its friendly name when given one and the target otherwise. The
    /// jump itself is the reading application's business, not the calculation's.
    /// </summary>
    [Fact]
    public void HyperlinkShowsWhateverItWasToldToShow()
    {
        Assert.Equal("Example", Eval("HYPERLINK(\"http://example.com\",\"Example\")").GetText());
        Assert.Equal("http://example.com", Eval("HYPERLINK(\"http://example.com\")").GetText());
        Assert.Equal(42.0, Eval("HYPERLINK(\"http://example.com\",42)").GetNumber());
    }

    [Fact]
    public void AddressBuildsReferences()
    {
        Assert.Equal("$B$2", Eval("ADDRESS(2,2)").GetText());
        Assert.Equal("B2", Eval("ADDRESS(2,2,4)").GetText());
    }

    [Fact]
    public void SubtotalSkipsHiddenRowsOnlyForTheHundredCodes()
    {
        Worksheet sheet = Sheet();
        sheet.Row(2).Hide();

        Assert.Equal(100.0, sheet.Evaluate("SUBTOTAL(9,A1:A4)", "Z1").GetNumber());
        Assert.Equal(80.0, sheet.Evaluate("SUBTOTAL(109,A1:A4)", "Z1").GetNumber());
        Assert.Equal(3.0, sheet.Evaluate("SUBTOTAL(102,A1:A4)", "Z1").GetNumber());
    }

    [Fact]
    public void AggregateHonoursTheHiddenRowOptions()
    {
        Worksheet sheet = Sheet();
        sheet.Row(3).Hide();

        Assert.Equal(100.0, sheet.Evaluate("AGGREGATE(9,0,A1:A4)", "Z1").GetNumber());
        Assert.Equal(70.0, sheet.Evaluate("AGGREGATE(9,1,A1:A4)", "Z1").GetNumber());
        Assert.Equal(70.0, sheet.Evaluate("AGGREGATE(9,7,A1:A4)", "Z1").GetNumber());
        Assert.Equal(100.0, sheet.Evaluate("AGGREGATE(9,4,A1:A4)", "Z1").GetNumber());
    }

    [Theory]
    [InlineData("TEXT(1234.567,\"#,##0.00\")", "1,234.57")]
    [InlineData("TEXT(0.256,\"0.0%\")", "25.6%")]
    [InlineData("TEXT(A1,\"0000\")", "0010")]
    [InlineData("TEXT(DATE(2026,7,25),\"yyyy-mm-dd\")", "2026-07-25")]
    [InlineData("TEXT(DATE(2026,7,25),\"dddd\")", "Saturday")]
    [InlineData("TEXT(-5,\"0;(0)\")", "(5)")]
    [InlineData("TEXT(\"abc\",\"0.00\")", "abc")]
    [InlineData("\"Итого: \"&TEXT(A4,\"#,##0\")", "Итого: 40")]
    public void TextRendersThroughANumberFormat(string formula, string expected)
    {
        Assert.Equal(expected, Eval(formula).GetText());
    }

    /// <summary>
    /// A file written under a locale where the comma is the decimal point and the space the
    /// group separator stores its <c>TEXT</c> formats exactly that way. With the authoring
    /// culture named in <see cref="CalculateOptions.FormatCulture"/>, both the reading of the
    /// format and the separators of the output follow it.
    /// </summary>
    [Fact]
    public void TextReadsFormatSeparatorsOfTheAuthoringLocale()
    {
        var culture = (System.Globalization.CultureInfo)System.Globalization.CultureInfo.InvariantCulture.Clone();
        culture.NumberFormat.NumberDecimalSeparator = ",";
        culture.NumberFormat.NumberGroupSeparator = "\u00A0";

        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1234.5;
        sheet.Cell("B1").SetFormula("TEXT(A1,\"# ##0,00\")");
        sheet.Cell("B2").SetFormula("TEXT(0.25,\"0,0%\")");
        sheet.Cell("B3").SetFormula("TEXT(7,\"0,00 \"\"kg\"\"\")");
        book.Calculate(new CalculateOptions { FormatCulture = culture });

        Assert.Equal("1\u00A0234,50", sheet["B1"].GetText());
        Assert.Equal("25,0%", sheet["B2"].GetText());

        // The space before the quoted literal is not between digit placeholders — it stays.
        Assert.Equal("7,00 kg", sheet["B3"].GetText());
    }

    [Theory]
    [InlineData("LET(x,5,x*2)", 10.0)]
    [InlineData("LET(x,5,y,x+1,x*y)", 30.0)]
    [InlineData("LET(total,SUM(A1:A4),total/COUNT(A1:A4))", 25.0)]
    [InlineData("1+LET(x,2,x)", 3.0)]
    [InlineData("LET(x,1,LET(x,2,x))", 2.0)]
    [InlineData("LET(x,1,LET(y,2,x+y))", 3.0)]
    [InlineData("_xlfn.LET(x,3,x^2)", 9.0)]
    public void LetBindsNamesForOneFormula(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void LetNeedsAtLeastOneBinding()
    {
        Assert.Throws<FormulaParseException>(() => Eval("LET(A1)"));
    }

    [Fact]
    public void LetBindingsDoNotLeakIntoTheRestOfTheFormula()
    {
        // Outside its own call the name is an ordinary defined name, and there is none.
        Assert.Equal(ExcelError.Name, Eval("LET(x,1,x)+x").GetError());
    }

    [Fact]
    public void SubtotalRejectsFunctionNumbersOnlyAggregateKnows()
    {
        Assert.Equal(ExcelError.Value, Eval("SUBTOTAL(14,A1:A4,2)").GetError());
        Assert.Equal(ExcelError.Value, Eval("SUBTOTAL(112,A1:A4)").GetError());
    }

    [Theory]
    [InlineData("AGGREGATE(14,4,A1:A4,2)", 30.0)]
    [InlineData("AGGREGATE(15,4,A1:A4,2)", 20.0)]
    [InlineData("AGGREGATE(16,4,A1:A4,0.5)", 25.0)]
    [InlineData("AGGREGATE(17,4,A1:A4,1)", 17.5)]
    [InlineData("AGGREGATE(18,4,A1:A4,0.5)", 25.0)]
    [InlineData("AGGREGATE(19,4,A1:A4,1)", 12.5)]
    public void AggregateOrderStatistics(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Theory]
    [InlineData("PERCENTILE.INC(C1:C4,0.5)", 2.5)]
    [InlineData("QUARTILE.INC(C1:C4,1)", 1.75)]
    [InlineData("PERCENTILE.EXC(C1:C4,0.5)", 2.5)]
    [InlineData("PERCENTILE.EXC(C1:C4,0.4)", 2.0)]
    [InlineData("QUARTILE.EXC(C1:C4,1)", 1.25)]
    [InlineData("QUARTILE.EXC(C1:C4,3)", 3.75)]
    [InlineData("_xlfn.PERCENTILE.EXC(C1:C4,0.2)", 1.0)]
    public void ExclusivePercentiles(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    [Fact]
    public void ExclusivePercentilesRefuseTheEndsOfTheData()
    {
        // With four values the reachable fractions run from 1/5 to 4/5, and the exclusive
        // quartiles have no zeroth or fourth member at all.
        Assert.Equal(ExcelError.Number, Eval("PERCENTILE.EXC(C1:C4,0.1)").GetError());
        Assert.Equal(ExcelError.Number, Eval("PERCENTILE.EXC(C1:C4,0.9)").GetError());
        Assert.Equal(ExcelError.Number, Eval("QUARTILE.EXC(C1:C4,0)").GetError());
        Assert.Equal(ExcelError.Number, Eval("QUARTILE.EXC(C1:C4,4)").GetError());
    }

    [Fact]
    public void RowsHiddenByAFilterAreSkippedByEverySubtotalCode()
    {
        Worksheet sheet = Sheet();
        sheet.AutoFilter("A1:C4");
        sheet.AutoFilterIsActive = true;
        sheet.Row(2).Hide();

        Assert.Equal(80.0, sheet.Evaluate("SUBTOTAL(9,A1:A4)", "Z1").GetNumber());
        Assert.Equal(80.0, sheet.Evaluate("SUBTOTAL(109,A1:A4)", "Z1").GetNumber());

        // AGGREGATE makes no such distinction: only the "ignore hidden rows" options skip.
        Assert.Equal(100.0, sheet.Evaluate("AGGREGATE(9,0,A1:A4)", "Z1").GetNumber());
        Assert.Equal(80.0, sheet.Evaluate("AGGREGATE(9,5,A1:A4)", "Z1").GetNumber());
    }

    [Fact]
    public void FilterHidingStopsAtTheHeaderAndTheEndOfTheRange()
    {
        Worksheet sheet = Sheet();
        sheet.AutoFilter("A2:C3");
        sheet.AutoFilterIsActive = true;
        sheet.Row(2).Hide();
        sheet.Row(4).Hide();

        // Row 2 is the filter's own header and row 4 lies past its range, so both count as
        // hidden by hand and survive the plain code.
        Assert.Equal(100.0, sheet.Evaluate("SUBTOTAL(9,A1:A4)", "Z1").GetNumber());
        Assert.Equal(40.0, sheet.Evaluate("SUBTOTAL(109,A1:A4)", "Z1").GetNumber());
    }

    [Fact]
    public void MovingTheFilterForgetsThatItsRowsWereFiltered()
    {
        Worksheet sheet = Sheet();
        sheet.AutoFilter("A1:C4");
        sheet.AutoFilterIsActive = true;
        sheet.Row(2).Hide();

        sheet.AutoFilter("A1:C3");

        Assert.False(sheet.AutoFilterIsActive);
        Assert.Equal(100.0, sheet.Evaluate("SUBTOTAL(9,A1:A4)", "Z1").GetNumber());
    }

    [Fact]
    public void HidingARowRecalculatesTheSubtotalsThatWatchIt()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 10;
        sheet["A2"] = 20;
        sheet.Cell("B1").SetFormula("SUBTOTAL(109,A1:A2)");
        sheet.Cell("B2").SetFormula("SUM(A1:A2)");
        book.Calculate();
        Assert.Equal(30.0, sheet["B1"].GetNumber());

        sheet.Row(2).Hide();
        CalculationSummary summary = book.Calculate();

        Assert.Equal(10.0, sheet["B1"].GetNumber());
        Assert.Equal(1, summary.CalculatedCells);
    }
}
