using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// A second <c>Calculate</c> has to answer the same as a fresh workbook would. These cases change
/// something between passes that the graph has to hear about: a name, a spill, the sheet list.
/// </summary>
public sealed class RecalculationContractTests
{
    private static (Workbook Book, Worksheet Sheet) NewBook(string name = "S")
    {
        var book = Workbook.Create();
        return (book, book.Sheets.Add(name));
    }

    // Defined names -----------------------------------------------------------------------------

    [Fact]
    public void FirstCalculationOrdersAFormulaBehindTheNameItUses()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Rate", "S!$B$1");
        sheet.Cell("A1").SetFormula("Rate*100");
        sheet.Cell("B1").SetFormula("1+1");

        book.Calculate();

        Assert.Equal(200.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void EditingTheCellANamePointsAtRecalculatesItsUsers()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["B1"] = 1;
        book.Names.Add("Rate", "S!$B$1");
        sheet.Cell("A1").SetFormula("Rate*100");
        book.Calculate();
        Assert.Equal(100.0, sheet["A1"].GetNumber());

        sheet["B1"] = 2;
        CalculationSummary summary = book.Calculate();

        Assert.Equal(200.0, sheet["A1"].GetNumber());
        Assert.Equal(1, summary.CalculatedCells);
    }

    [Fact]
    public void NamesPointingAtRangesRecalculateTheirUsers()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["B1"] = 1;
        sheet["B2"] = 2;
        book.Names.Add("Values", "S!$B$1:$B$2");
        sheet.Cell("A1").SetFormula("SUM(Values)");
        book.Calculate();

        sheet["B2"] = 40;
        book.Calculate();

        Assert.Equal(41.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void ReplacingAConstantNameRecalculatesItsUsers()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Rate", "1");
        sheet.Cell("A1").SetFormula("Rate*100");
        book.Calculate();
        Assert.Equal(100.0, sheet["A1"].GetNumber());

        book.Names.Add("Rate", "2");
        book.Calculate();

        Assert.Equal(200.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void RedirectingANameSubscribesToItsNewSource()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["B1"] = 1;
        sheet["B2"] = 5;
        book.Names.Add("Rate", "S!$B$1");
        sheet.Cell("A1").SetFormula("Rate*100");
        book.Calculate();

        book.Names.Add("Rate", "S!$B$2");
        book.Calculate();
        Assert.Equal(500.0, sheet["A1"].GetNumber());

        sheet["B2"] = 6;
        book.Calculate();
        Assert.Equal(600.0, sheet["A1"].GetNumber());

        sheet["B1"] = 99;
        Assert.Equal(0, book.Calculate().CalculatedCells);
    }

    [Fact]
    public void RemovingANameTurnsItsUsersIntoNameErrors()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Rate", "2");
        sheet.Cell("A1").SetFormula("Rate*100");
        book.Calculate();

        Assert.True(book.Names.Remove("Rate"));
        book.Calculate();

        Assert.Equal(CellValue.Error(ExcelError.Name), sheet["A1"]);
    }

    [Fact]
    public void AddingANameFixesFormulasThatWereWaitingForIt()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet.Cell("A1").SetFormula("Later*2");
        book.Calculate();
        Assert.Equal(CellValue.Error(ExcelError.Name), sheet["A1"]);

        book.Names.Add("Later", "21");
        book.Calculate();

        Assert.Equal(42.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void NamesFollowOtherNamesToTheirSource()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["C1"] = 3;
        book.Names.Add("Outer", "Inner*2");
        book.Names.Add("Inner", "S!$C$1");
        sheet.Cell("A1").SetFormula("Outer");
        book.Calculate();
        Assert.Equal(6.0, sheet["A1"].GetNumber());

        sheet["C1"] = 4;
        book.Calculate();

        Assert.Equal(8.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void ASheetScopedNameShadowsTheWorkbookNameOnlyOnItsOwnSheet()
    {
        var book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");
        first["B1"] = 1;
        first["B2"] = 2;
        book.Names.Add("Rate", "First!$B$1");
        first.Cell("A1").SetFormula("Rate");
        second.Cell("A1").SetFormula("Rate");
        book.Calculate();

        book.Names.Add("Rate", "First!$B$2", first);
        book.Calculate();
        Assert.Equal(2.0, first["A1"].GetNumber());
        Assert.Equal(1.0, second["A1"].GetNumber());

        first["B2"] = 20;
        first["B1"] = 10;
        book.Calculate();
        Assert.Equal(20.0, first["A1"].GetNumber());
        Assert.Equal(10.0, second["A1"].GetNumber());
    }

    [Fact]
    public void AVolatileNameMakesItsUsersVolatile()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Stamp", "NOW()");
        sheet.Cell("A1").SetFormula("Stamp");
        book.Calculate();

        Assert.Equal(1, book.Calculate().CalculatedCells);
    }

    [Fact]
    public void ANameThatLeadsBackToItsUserIsACircularReference()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Loop", "S!$A$1");
        sheet.Cell("A1").SetFormula("Loop+1");

        CalculationSummary summary = book.Calculate();

        Assert.True(summary.HasCircularReferences);
        Assert.False(book.IsFullyCalculated);
    }

    [Fact]
    public void NamesReferringToEachOtherDoNotHangTheGraph()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        book.Names.Add("Ping", "Pong");
        book.Names.Add("Pong", "Ping");
        sheet.Cell("A1").SetFormula("Ping");

        book.Calculate();

        Assert.True(sheet["A1"].IsError);
    }

    // Spills --------------------------------------------------------------------------------------

    private static (Workbook Book, Worksheet Sheet) SpilledBook()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["B1"] = 1;
        sheet.Cell("A1").SetFormula("SEQUENCE(3,1,B1)");
        return (book, sheet);
    }

    [Fact]
    public void ReadersOfASpilledCellFollowNewValuesInTheSameRectangle()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        sheet.Cell("C1").SetFormula("A3*10");
        book.Calculate();
        Assert.Equal(30.0, sheet["C1"].GetNumber());

        sheet["B1"] = 10;
        book.Calculate();

        Assert.Equal(12.0, sheet["A3"].GetNumber());
        Assert.Equal(120.0, sheet["C1"].GetNumber());
    }

    [Fact]
    public void ReadersOfASpilledRangeFollowNewValuesInTheSameRectangle()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        sheet.Cell("C1").SetFormula("SUM(A2:A3)");
        book.Calculate();
        Assert.Equal(5.0, sheet["C1"].GetNumber());

        sheet["B1"] = 10;
        book.Calculate();

        Assert.Equal(23.0, sheet["C1"].GetNumber());
    }

    [Fact]
    public void AnUnchangedSpillDoesNotRerunItsReaders()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        sheet.Cell("C1").SetFormula("A3*10");
        book.Calculate();

        // Re-evaluating the anchor with the same result must not cascade into a follow-up round.
        sheet["B1"] = 1;

        Assert.Equal(1, book.Calculate().CalculatedCells);
    }

    [Fact]
    public void ChainedSpillsSettleWithinOnePass()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        sheet.Cell("C1").SetFormula("A3*2");
        sheet.Cell("D1").SetFormula("SEQUENCE(2,1,C1)");
        sheet.Cell("F1").SetFormula("D2+1");
        book.Calculate();
        Assert.Equal(8.0, sheet["F1"].GetNumber());

        sheet["B1"] = 10;
        book.Calculate();

        Assert.Equal(26.0, sheet["F1"].GetNumber());
    }

    [Fact]
    public void ReplacingTheAnchorWithAValueTakesTheSpillBack()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        sheet.Cell("C1").SetFormula("A3+1");
        book.Calculate();
        Assert.Equal(new RangeRef(new CellRef(1, 1), new CellRef(3, 1)), sheet.Cell("A1").SpillRange);

        sheet["A1"] = 42;
        book.Calculate();

        Assert.Equal(42.0, sheet["A1"].GetNumber());
        Assert.True(sheet["A2"].IsBlank);
        Assert.True(sheet["A3"].IsBlank);
        Assert.Null(sheet.Cell("A1").SpillRange);
        Assert.Equal(1.0, sheet["C1"].GetNumber());
    }

    [Fact]
    public void ClearingTheAnchorTakesTheSpillBack()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        book.Calculate();

        sheet.Cell("A1").Clear();
        book.Calculate();

        Assert.True(sheet["A2"].IsBlank);
        Assert.True(sheet["A3"].IsBlank);
        Assert.Null(sheet.Cell("A1").SpillRange);
    }

    [Fact]
    public void InsertingARowAboveTheAnchorMovesTheSpillWithIt()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        book.Calculate();

        sheet.InsertRows(1, 1);
        book.Calculate();

        Assert.Equal("SEQUENCE(3,1,B2)", sheet.Cell("A2").Formula);
        Assert.Equal(new RangeRef(new CellRef(2, 1), new CellRef(4, 1)), sheet.Cell("A2").SpillRange);
        Assert.Equal(3.0, sheet["A4"].GetNumber());
        Assert.True(sheet["A1"].IsBlank);
    }

    [Fact]
    public void InsertingARowInsideTheRectangleReflowsTheSpill()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        book.Calculate();

        sheet.InsertRows(2, 1);
        book.Calculate();

        Assert.Equal(new RangeRef(new CellRef(1, 1), new CellRef(3, 1)), sheet.Cell("A1").SpillRange);
        Assert.Equal(2.0, sheet["A2"].GetNumber());
        Assert.Equal(3.0, sheet["A3"].GetNumber());
        Assert.True(sheet["A4"].IsBlank);
    }

    [Fact]
    public void DeletingTheAnchorRowTakesTheRestOfTheSpillWithIt()
    {
        (Workbook book, Worksheet sheet) = SpilledBook();
        book.Calculate();

        sheet.DeleteRows(1, 1);
        book.Calculate();

        Assert.True(sheet["A1"].IsBlank);
        Assert.True(sheet["A2"].IsBlank);
        Assert.Empty(sheet.SpillTable);
    }

    // The sheet list -----------------------------------------------------------------------------

    [Fact]
    public void RemovingASheetTurnsReferencesToItIntoRefErrors()
    {
        var book = Workbook.Create();
        Worksheet data = book.Sheets.Add("Data");
        Worksheet report = book.Sheets.Add("Report");
        data["A1"] = 3;
        report.Cell("A1").SetFormula("Data!A1*2");
        book.Calculate();
        Assert.Equal(6.0, report["A1"].GetNumber());

        book.Sheets.Remove("Data");
        book.Calculate();

        Assert.Equal(CellValue.Error(ExcelError.Reference), report["A1"]);
    }

    [Fact]
    public void AddingAMissingSheetFixesFormulasThatReferToIt()
    {
        var book = Workbook.Create();
        Worksheet report = book.Sheets.Add("Report");
        report.Cell("A1").SetFormula("Future!A1*2");
        book.Calculate();
        Assert.True(report["A1"].IsError);

        Worksheet future = book.Sheets.Add("Future");
        future["A1"] = 21;
        book.Calculate();

        Assert.Equal(42.0, report["A1"].GetNumber());
    }

    [Fact]
    public void MovingASheetOutOfA3DSpanRecalculatesTheSpan()
    {
        var book = Workbook.Create();
        Worksheet q1 = book.Sheets.Add("Q1");
        Worksheet q2 = book.Sheets.Add("Q2");
        Worksheet q3 = book.Sheets.Add("Q3");
        Worksheet total = book.Sheets.Add("Total");
        q1["A1"] = 1;
        q2["A1"] = 2;
        q3["A1"] = 3;
        total.Cell("A1").SetFormula("SUM(Q1:Q3!A1)");
        book.Calculate();
        Assert.Equal(6.0, total["A1"].GetNumber());

        book.Sheets.Move(q2, 3);
        book.Calculate();

        Assert.Equal(4.0, total["A1"].GetNumber());

        q1["A1"] = 10;
        book.Calculate();
        Assert.Equal(13.0, total["A1"].GetNumber());
    }

    [Fact]
    public void ARenameLetsAWaitingFormulaFindTheSheet()
    {
        var book = Workbook.Create();
        Worksheet old = book.Sheets.Add("Old");
        Worksheet report = book.Sheets.Add("Report");
        old["A1"] = 5;
        report.Cell("A1").SetFormula("New!A1");
        book.Calculate();
        Assert.True(report["A1"].IsError);

        old.Name = "New";
        book.Calculate();

        Assert.Equal(5.0, report["A1"].GetNumber());
    }

    // Cancellation -------------------------------------------------------------------------------

    private static CalculateOptions Cancelled()
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        return new CalculateOptions { CancellationToken = source.Token };
    }

    [Fact]
    public void ACancelledFirstPassIsRepeatedInFull()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet.Cell("A1").SetFormula("2+2");

        Assert.Throws<OperationCanceledException>(() => book.Calculate(Cancelled()));
        Assert.False(book.IsFullyCalculated);

        CalculationSummary retry = book.Calculate();

        Assert.Equal(1, retry.CalculatedCells);
        Assert.Equal(4.0, sheet["A1"].GetNumber());
        Assert.True(book.IsFullyCalculated);
    }

    [Fact]
    public void ACancelledPartialPassKeepsItsWork()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["B1"] = 1;
        sheet.Cell("A1").SetFormula("B1*2");
        book.Calculate();

        sheet["B1"] = 5;
        Assert.Throws<OperationCanceledException>(() => book.Calculate(Cancelled()));
        Assert.False(book.IsFullyCalculated);

        CalculationSummary retry = book.Calculate();

        Assert.Equal(1, retry.CalculatedCells);
        Assert.Equal(10.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void CalculatedValueRecoversAfterACancelledPass()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet.Cell("A1").SetFormula("2+2");
        Assert.Throws<OperationCanceledException>(() => book.Calculate(Cancelled()));

        Assert.Equal(4.0, sheet.Cell("A1").CalculatedValue.GetNumber());
    }

    // Circular references through ranges ----------------------------------------------------------

    [Fact]
    public void ARangeContainingTheFormulaItselfIsCircular()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["A2"] = 1;
        sheet.Cell("A1").SetFormula("SUM(A1:A2)");

        CalculationSummary summary = book.Calculate();

        Assert.True(summary.HasCircularReferences);
        Assert.False(book.IsFullyCalculated);
        Assert.Equal(0.0, sheet["A1"].GetNumber());
    }

    [Fact]
    public void AWholeColumnContainingTheFormulaIsCircular()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["A1"] = 1;
        sheet.Cell("A2").SetFormula("SUM(A:A)");

        Assert.True(book.Calculate().HasCircularReferences);
    }

    [Fact]
    public void A3DSpanContainingTheFormulaIsCircular()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Q1")["A1"] = 1;
        Worksheet q2 = book.Sheets.Add("Q2");
        book.Sheets.Add("Q3")["A1"] = 3;
        q2.Cell("A1").SetFormula("SUM(Q1:Q3!A1)");

        Assert.True(book.Calculate().HasCircularReferences);
    }

    [Fact]
    public void TwoFormulasThroughARangeAreCircular()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet.Cell("A1").SetFormula("SUM(B1:B2)");
        sheet.Cell("B2").SetFormula("A1+1");

        Assert.True(book.Calculate().HasCircularReferences);
    }

    [Theory]
    [InlineData("A1", "ROW(A1)", 1.0)]
    [InlineData("A2", "ROWS(A1:A5)", 5.0)]
    [InlineData("B1", "COLUMN(B1)", 2.0)]
    [InlineData("B1", "COLUMNS(A1:C1)", 3.0)]
    public void FunctionsThatOnlyReadAnAddressAreNotCircular(string cell, string formula, double expected)
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet.Cell(cell).SetFormula(formula);

        CalculationSummary summary = book.Calculate();

        Assert.False(summary.HasCircularReferences);
        Assert.Equal(expected, sheet[cell].GetNumber());
    }

    [Fact]
    public void AnIndexOverItsOwnColumnIsNotCircular()
    {
        (Workbook book, Worksheet sheet) = NewBook();
        sheet["A1"] = 2;
        sheet["A2"] = 3;
        sheet.Cell("A3").SetFormula("INDEX(A:A,ROW()-1)");

        CalculationSummary summary = book.Calculate();

        Assert.False(summary.HasCircularReferences);
        Assert.Equal(3.0, sheet["A3"].GetNumber());
    }
}
