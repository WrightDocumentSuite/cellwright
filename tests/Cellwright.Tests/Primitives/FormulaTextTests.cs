using Xunit;

namespace Cellwright.Tests.Primitives;

/// <summary>
/// The textual rewriter behind sheet renaming and structural editing. What it must never do
/// is touch anything it did not recognize, so most of these check the parts left alone.
/// </summary>
public sealed class FormulaTextTests
{
    [Theory]
    [InlineData("Data!A1", "Sales!A1")]
    [InlineData("SUM(Data!A1:B9)*2", "SUM(Sales!A1:B9)*2")]
    [InlineData("data!A1+DATA!B2", "Sales!A1+Sales!B2")]
    [InlineData("'Data'!A1", "Sales!A1")]
    [InlineData("Data!A:A", "Sales!A:A")]
    [InlineData("Data!$1:$5", "Sales!$1:$5")]
    public void RenameSheet_RewritesEveryPrefix(string formula, string expected)
    {
        Assert.Equal(expected, FormulaText.RenameSheet(formula, "Data", "Sales"));
    }

    [Theory]
    [InlineData("Other!A1")]
    [InlineData("\"Data!A1\"")]
    [InlineData("DataRate*2")]
    [InlineData("SUM(A1:A9)")]
    public void RenameSheet_LeavesUnrelatedTextAlone(string formula)
    {
        Assert.Same(formula, FormulaText.RenameSheet(formula, "Data", "Sales"));
    }

    [Fact]
    public void RenameSheet_QuotesWhatExcelWouldQuote()
    {
        Assert.Equal("'Q1 2026'!A1", FormulaText.RenameSheet("Data!A1", "Data", "Q1 2026"));
        Assert.Equal("'It''s'!A1", FormulaText.RenameSheet("Data!A1", "Data", "It's"));
        Assert.Equal("'2026'!A1", FormulaText.RenameSheet("Data!A1", "Data", "2026"));
    }

    [Theory]
    [InlineData("SUM(Data:Other!A1)", "SUM(Sales:Other!A1)")]
    [InlineData("SUM(Other:Data!A1:B2)", "SUM(Other:Sales!A1:B2)")]
    [InlineData("SUM('First:Data'!A1)", "SUM(First:Sales!A1)")]
    [InlineData("SUM('Q 1:Data'!A1)", "SUM('Q 1:Sales'!A1)")]
    public void RenameSheet_FollowsSpanEndpoints(string formula, string expected)
    {
        Assert.Equal(expected, FormulaText.RenameSheet(formula, "Data", "Sales"));
    }

    [Fact]
    public void Shift_LeavesSpansAndOtherBooksAlone()
    {
        var shift = new ReferenceShift("Data", Rows: true, Index: 1, Count: 5);

        // A 3-D span names one rectangle over many sheets; no single shift fits them all.
        Assert.Same("SUM(Data:Other!A5)", FormulaText.Shift("SUM(Data:Other!A5)", shift, "Data"));

        // A reference into another workbook, bracket-index or full-path form, is not ours to move.
        Assert.Same("[1]Data!A5*2", FormulaText.Shift("[1]Data!A5*2", shift, "Data"));
        Assert.Same(
            "'[C:\\Reports\\Last year.xls]Data'!$D:$I",
            FormulaText.Shift("'[C:\\Reports\\Last year.xls]Data'!$D:$I", shift, "Data"));
    }

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("A5", "A7")]
    [InlineData("$A$5", "$A$7")]
    [InlineData("SUM(A5:A9)", "SUM(A7:A11)")]
    [InlineData("SUM(A1:A9)", "SUM(A1:A11)")]
    [InlineData("SUM(A1:A2)", "SUM(A1:A2)")]
    [InlineData("SUM(5:9)", "SUM(7:11)")]
    [InlineData("SUM(A:A)", "SUM(A:A)")]
    public void Shift_InsertingRowsPushesAndStretches(string formula, string expected)
    {
        var shift = new ReferenceShift("S", Rows: true, Index: 3, Count: 2);
        Assert.Equal(expected, FormulaText.Shift(formula, shift, "S"));
    }

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("A9", "A7")]
    [InlineData("A3", "#REF!")]
    [InlineData("SUM(A1:A9)", "SUM(A1:A7)")]
    [InlineData("SUM(A3:A4)", "SUM(#REF!)")]
    [InlineData("SUM(A2:A9)", "SUM(A2:A7)")]
    [InlineData("SUM(A3:A9)", "SUM(A3:A7)")]
    public void Shift_DeletingRowsPullsBackAndBreaks(string formula, string expected)
    {
        var shift = new ReferenceShift("S", Rows: true, Index: 3, Count: -2);
        Assert.Equal(expected, FormulaText.Shift(formula, shift, "S"));
    }

    [Theory]
    [InlineData("B1", "D1")]
    [InlineData("SUM(A1:B1)", "SUM(A1:D1)")]
    [InlineData("SUM(B:C)", "SUM(D:E)")]
    [InlineData("SUM(1:5)", "SUM(1:5)")]
    public void Shift_InsertingColumnsMovesTheColumnAxisOnly(string formula, string expected)
    {
        var shift = new ReferenceShift("S", Rows: false, Index: 2, Count: 2);
        Assert.Equal(expected, FormulaText.Shift(formula, shift, "S"));
    }

    [Fact]
    public void Shift_OnlyTouchesTheSheetThatChanged()
    {
        var shift = new ReferenceShift("Data", Rows: true, Index: 1, Count: 1);

        Assert.Equal("Data!A2+Sales!A1+A1", FormulaText.Shift("Data!A1+Sales!A1+A1", shift, "Sales"));
        Assert.Equal("Data!A2+Sales!A1+A2", FormulaText.Shift("Data!A1+Sales!A1+A1", shift, "Data"));
    }

    [Fact]
    public void Shift_LeavesStringsFunctionNamesAndErrorsAlone()
    {
        var shift = new ReferenceShift("S", Rows: true, Index: 1, Count: 1);

        Assert.Equal("IF(A2=\"A1\",LOG10(A2),#REF!)", FormulaText.Shift("IF(A1=\"A1\",LOG10(A1),#REF!)", shift, "S"));
        Assert.Equal("TaxRate*A2", FormulaText.Shift("TaxRate*A1", shift, "S"));
    }

    [Fact]
    public void Shift_KeepsSpacingAndCasingOfWhatItDoesNotRewrite()
    {
        var shift = new ReferenceShift("S", Rows: true, Index: 10, Count: 1);

        Assert.Same("sum( A1 : A2 )", FormulaText.Shift("sum( A1 : A2 )", shift, "S"));
    }

    [Fact]
    public void Shift_StopsAtTheEdgeOfTheGrid()
    {
        var shift = new ReferenceShift("S", Rows: true, Index: 1, Count: 5);

        Assert.Equal("A1048576", FormulaText.Shift("A1048574", shift, "S"));
    }

    [Fact]
    public void Translate_MovesOnlyRelativeAxesFromTheSourceCell()
    {
        const string formula =
            "A1+$A1+A$1+$A$1+A1:B2+A:C+$A:$C+1:2+$1:2+Other!A1+SUM(Data:Other!A1)+\"A1\"+LOG10(A1)";
        const string expected =
            "C4+$A4+C$1+$A$1+C4:D5+C:E+$A:$C+4:5+$1:5+Other!C4+SUM(Data:Other!C4)+\"A1\"+LOG10(C4)";

        Assert.Equal(expected, FormulaText.Translate(formula, CellRef.Parse("B2"), CellRef.Parse("D5")));
    }

    [Fact]
    public void Translate_TurnsReferencesCopiedOutsideTheGridIntoRefErrors()
    {
        Assert.Equal(
            "#REF!+A1",
            FormulaText.Translate("A1+B2", CellRef.Parse("B2"), CellRef.Parse("A1")));
    }

    [Fact]
    public void Translate_DoesNotTreatStructuredReferenceColumnsAsCells()
    {
        const string formula = "Table1[A1]+[@B2]+Table1[[#Headers],[C3]]+$A1";

        Assert.Equal(
            "Table1[A1]+[@B2]+Table1[[#Headers],[C3]]+$A4",
            FormulaText.Translate(formula, CellRef.Parse("B2"), CellRef.Parse("D5")));
        Assert.Equal(
            "Table[A']B2]+$A4",
            FormulaText.Translate("Table[A']B2]+$A1", CellRef.Parse("B2"), CellRef.Parse("D5")));
    }

    [Fact]
    public void Translate_ReturnsTheOriginalTextWhenNothingMoves()
    {
        const string formula = "$A$1+\"B2\"";

        Assert.Same(formula, FormulaText.Translate(formula, CellRef.Parse("A1"), CellRef.Parse("C3")));
        Assert.Same("A1", FormulaText.Translate("A1", CellRef.Parse("B2"), CellRef.Parse("B2")));
    }
}
