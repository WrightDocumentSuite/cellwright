using Cellwright.Formulas.Parsing;
using Xunit;

namespace Cellwright.Formulas.Tests;

public sealed class ParserTests
{
    private static string RoundTrip(string formula, Workbook? workbook = null) =>
        FormulaRenderer.Render(FormulaParser.Parse(formula, workbook), workbook);

    [Theory]
    [InlineData("1+2", "1+2")]
    [InlineData("=1+2", "1+2")]
    [InlineData("1+2*3", "1+2*3")]
    [InlineData("(1+2)*3", "(1+2)*3")]
    [InlineData("2^3^2", "2^3^2")]
    [InlineData("-2^2", "-2^2")]
    [InlineData("A1", "A1")]
    [InlineData("$A$1", "$A$1")]
    [InlineData("A1:B5", "A1:B5")]
    [InlineData("A:A", "A:A")]
    [InlineData("1:3", "1:3")]
    [InlineData("SUM(A1:A5)", "SUM(A1:A5)")]
    [InlineData("SUM(A1,B1,C1)", "SUM(A1,B1,C1)")]
    [InlineData("IF(A1>0,\"yes\",\"no\")", "IF(A1>0,\"yes\",\"no\")")]
    [InlineData("50%", "50%")]
    [InlineData("\"a\"&\"b\"", "\"a\"&\"b\"")]
    [InlineData("{1,2;3,4}", "{1,2;3,4}")]
    [InlineData("A1<>B1", "A1<>B1")]
    [InlineData("#DIV/0!", "#DIV/0!")]
    [InlineData("SUM(A1:A3 A2:A5)", "SUM(A1:A3 A2:A5)")]
    [InlineData("LET(x,1,x+1)", "LET(x,1,x+1)")]
    [InlineData("LET(x,A1,y,x*2,x+y)", "LET(x,A1,y,x*2,x+y)")]
    [InlineData("1+LET(x,2,x)", "1+LET(x,2,x)")]
    [InlineData("LET(x,1,LET(y,2,x+y))", "LET(x,1,LET(y,2,x+y))")]
    public void ParseAndRender_RoundTrips(string formula, string expected)
    {
        Assert.Equal(expected, RoundTrip(formula));
    }

    [Fact]
    public void RenderIsStable_WhenParsedAgain()
    {
        string[] formulas =
        [
            "SUM(A1:A5)*2+MAX(B1:B9)", "IF(AND(A1>1,B1<2),\"x\",\"y\")", "-A1%", "(A1+B1)/(C1-D1)",
            "VLOOKUP($A2,Data!$A$1:$D$99,3,FALSE)", "SUMPRODUCT((A1:A9>3)*(B1:B9))",
        ];

        var book = Workbook.Create();
        book.Sheets.Add("Sheet1");
        book.Sheets.Add("Data");

        foreach (string formula in formulas)
        {
            string once = RoundTrip(formula, book);
            string twice = RoundTrip(once, book);
            Assert.Equal(once, twice);
        }
    }

    [Fact]
    public void SheetQualifiedReferences_KeepQuotingWhenNeeded()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Data");
        book.Sheets.Add("My Sheet");

        Assert.Equal("Data!A1", RoundTrip("Data!A1", book));
        Assert.Equal("'My Sheet'!A1:B2", RoundTrip("'My Sheet'!A1:B2", book));
    }

    [Fact]
    public void RenamingASheet_UpdatesRenderedFormulas()
    {
        var book = Workbook.Create();
        Worksheet data = book.Sheets.Add("Data");
        FormulaProgram program = FormulaParser.Parse("Data!A1*2", book);

        data.Name = "Renamed";

        Assert.Equal("Renamed!A1*2", FormulaRenderer.Render(program, book));
    }

    [Theory]
    [InlineData("A1", 1, 0, "A2")]
    [InlineData("A1", 0, 1, "B1")]
    [InlineData("$A1", 1, 1, "$A2")]
    [InlineData("A$1", 1, 1, "B$1")]
    [InlineData("$A$1", 5, 5, "$A$1")]
    [InlineData("SUM(A1:A3)", 2, 0, "SUM(A3:A5)")]
    public void Shifting_MovesOnlyRelativeParts(string formula, int rowDelta, int columnDelta, string expected)
    {
        FormulaProgram program = FormulaParser.Parse(formula);
        Assert.Equal(expected, FormulaRenderer.Render(program, null, rowDelta, columnDelta));
    }

    [Theory]
    [InlineData("LOG10(100)")]
    [InlineData("SUM(A1)")]
    public void IdentifiersThatLookLikeCells_AreFunctions(string formula)
    {
        FormulaProgram program = FormulaParser.Parse(formula);
        Assert.Contains(program.Instructions, i => i.Code == OpCode.Call);
    }

    [Fact]
    public void OutOfRangeIdentifiers_StayNames()
    {
        // XFE is past the last column, so Excel reads XFE1 as a defined name.
        FormulaProgram program = FormulaParser.Parse("XFE1");
        Assert.Contains(program.Instructions, i => i.Code == OpCode.PushName);
    }

    [Theory]
    [InlineData("")]
    [InlineData("1+")]
    [InlineData("(1")]
    [InlineData("SUM(")]
    [InlineData("\"unterminated")]
    [InlineData("{1,2;3}")]
    public void MalformedFormulas_ThrowParseException(string formula)
    {
        Assert.Throws<FormulaParseException>(() => FormulaParser.Parse(formula));
    }

    [Fact]
    public void IntersectingNonReferences_ParsesButEvaluatesToValueError()
    {
        // "1 2" is the intersection operator applied to numbers: syntactically fine,
        // semantically an error, which is how Excel reports it too.
        var book = Workbook.Create();
        book.Sheets.Add("S");
        Assert.Equal(ExcelError.Value, book.Evaluate("1 2").GetError());
    }

    [Fact]
    public void DeeplyNestedInput_FailsCleanlyInsteadOfOverflowing()
    {
        string formula = new string('(', 500) + "1" + new string(')', 500);
        FormulaParseException exception = Assert.Throws<FormulaParseException>(() => FormulaParser.Parse(formula));
        Assert.Contains("nests deeper", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WrongArgumentCount_IsRejectedAtParseTime()
    {
        Assert.Throws<FormulaParseException>(() => FormulaParser.Parse("LEFT()"));
        Assert.Throws<FormulaParseException>(() => FormulaParser.Parse("PI(1)"));
    }

    [Fact]
    public void UnknownFunction_ParsesAndBecomesNameErrorLater()
    {
        FormulaProgram program = FormulaParser.Parse("NOSUCHFUNC(1,2)");
        Assert.Contains(program.Instructions, i => i.Code == OpCode.Call);
        Assert.Null(program.Functions[0]);
    }

    [Fact]
    public void VolatileFunctions_MarkTheProgram()
    {
        Assert.True(FormulaParser.Parse("NOW()").IsVolatile);
        Assert.True(FormulaParser.Parse("A1+RAND()").IsVolatile);
        Assert.False(FormulaParser.Parse("A1+1").IsVolatile);
    }

    [Fact]
    public void FutureFunctionPrefix_IsIgnored()
    {
        FormulaProgram program = FormulaParser.Parse("_xlfn.XLOOKUP(A1,B:B,C:C)");
        Assert.NotNull(program.Functions[0]);
        Assert.Equal("XLOOKUP", program.Functions[0]!.Name);
    }

    [Fact]
    public void RepeatedOperands_AreDeduplicatedInThePool()
    {
        FormulaProgram program = FormulaParser.Parse("A1+A1+A1");
        Assert.Single(program.References);
    }
}
