using Cellwright.Formulas;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Dynamic array spill: an array result occupies the cells beside its formula, retreats when it
/// shrinks, and turns to <c>#SPILL!</c> whole when anything stands in its way.
/// </summary>
public sealed class SpillTests
{
    private static Worksheet Sheet(out Workbook book)
    {
        book = Workbook.Create();
        return book.Sheets.Add("Sheet1");
    }

    [Fact]
    public void AnArrayOccupiesItsRectangle()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");

        book.Calculate();

        Assert.Equal(1d, sheet["A1"].GetNumber());
        Assert.Equal(2d, sheet["A2"].GetNumber());
        Assert.Equal(3d, sheet["A3"].GetNumber());
    }

    [Fact]
    public void ATwoDimensionalArraySpillsBothWays()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("B2").SetFormula("SEQUENCE(2,3)");

        book.Calculate();

        Assert.Equal(1d, sheet["B2"].GetNumber());
        Assert.Equal(3d, sheet["D2"].GetNumber());
        Assert.Equal(4d, sheet["B3"].GetNumber());
        Assert.Equal(6d, sheet["D3"].GetNumber());
    }

    [Fact]
    public void TheAnchorKnowsItsRangeAndTheNeighboursKnowTheyAreSpilled()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");

        book.Calculate();

        Assert.Equal(RangeRef.Parse("A1:A3"), sheet.Cell("A1").SpillRange);
        Assert.False(sheet.Cell("A1").IsSpilled);
        Assert.True(sheet.Cell("A2").IsSpilled);
        Assert.Null(sheet.Cell("A2").SpillRange);
        Assert.Null(sheet.Cell("A2").Formula);
    }

    [Fact]
    public void AValueInTheWayMakesTheWholeFormulaSpillError()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        sheet["A3"] = "in the way";

        book.Calculate();

        Assert.Equal(ExcelError.Spill, sheet["A1"].GetError());

        // Nothing of the array is shown: a truncated array would read as the answer.
        Assert.True(sheet["A2"].IsBlank);
        Assert.Equal("in the way", sheet["A3"].GetText());
        Assert.Null(sheet.Cell("A1").SpillRange);
    }

    [Fact]
    public void AMergeInTheWayBlocksTheSpill()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Merge(RangeRef.Parse("A2:B2"));
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");

        book.Calculate();

        Assert.Equal(ExcelError.Spill, sheet["A1"].GetError());
    }

    [Fact]
    public void AShrinkingArrayVacatesWhatItNoLongerCovers()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["B1"] = 4d;
        sheet.Cell("A1").SetFormula("SEQUENCE(B1)");
        book.Calculate();
        Assert.Equal(4d, sheet["A4"].GetNumber());

        sheet["B1"] = 2d;
        book.Calculate();

        Assert.Equal(2d, sheet["A2"].GetNumber());
        Assert.True(sheet["A3"].IsBlank);
        Assert.True(sheet["A4"].IsBlank);
        Assert.Equal(RangeRef.Parse("A1:A2"), sheet.Cell("A1").SpillRange);
    }

    [Fact]
    public void AFormulaThatStopsReturningAnArrayTakesItsSpillWithIt()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();

        sheet.Cell("A1").SetFormula("42");
        book.Calculate();

        Assert.Equal(42d, sheet["A1"].GetNumber());
        Assert.True(sheet["A2"].IsBlank);
        Assert.True(sheet["A3"].IsBlank);
        Assert.Null(sheet.Cell("A1").SpillRange);
    }

    [Fact]
    public void RemovingTheBlockerLetsTheArrayThrough()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["A3"] = "in the way";
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();
        Assert.Equal(ExcelError.Spill, sheet["A1"].GetError());

        sheet.Cell("A3").Clear();
        book.Calculate();

        Assert.Equal(1d, sheet["A1"].GetNumber());
        Assert.Equal(3d, sheet["A3"].GetNumber());
    }

    [Fact]
    public void TypingIntoASpilledCellBreaksTheSpill()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();
        Assert.Equal(2d, sheet["A2"].GetNumber());

        sheet["A2"] = "mine now";
        book.Calculate();

        Assert.Equal(ExcelError.Spill, sheet["A1"].GetError());
        Assert.Equal("mine now", sheet["A2"].GetText());
    }

    [Fact]
    public void AFormulaReadingASpilledCellSeesTheSpilledValue()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        sheet.Cell("C1").SetFormula("A3*10");

        book.Calculate();

        Assert.Equal(30d, sheet["C1"].GetNumber());
    }

    [Fact]
    public void AFormulaSummingOverASpillSeesAllOfIt()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(4)");
        sheet.Cell("C1").SetFormula("SUM(A1:A4)");

        book.Calculate();

        Assert.Equal(10d, sheet["C1"].GetNumber());
    }

    [Fact]
    public void TwoSpillsLiveSideBySide()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        sheet.Cell("B1").SetFormula("SEQUENCE(2)");

        book.Calculate();

        Assert.Equal(3d, sheet["A3"].GetNumber());
        Assert.Equal(2d, sheet["B2"].GetNumber());
    }

    [Fact]
    public void TwoSpillsCannotClaimTheSameCell()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(1,3)");    // Across: A1..C1.
        sheet.Cell("C1").SetFormula("SEQUENCE(3)");      // Down: C1..C3 — collides at C1.

        book.Calculate();

        // One of the two spills and the other reports; which one wins follows evaluation
        // order, but they must never both write.
        int errors = 0;
        if (sheet["A1"].Kind == CellValueKind.Error)
            errors++;
        if (sheet["C1"].Kind == CellValueKind.Error && sheet.Cell("C1").SpillRange is null)
            errors++;

        Assert.Equal(1, errors);
    }

    [Fact]
    public void AnArrayFallingOffTheGridIsASpillError()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1048575").SetFormula("SEQUENCE(3)");

        book.Calculate();

        Assert.Equal(ExcelError.Spill, sheet["A1048575"].GetError());
    }

    [Fact]
    public void ErrorElementsSpillAsThemselves()
    {
        // An element-wise error belongs to its element: the rest of the array still spills,
        // which is what separates a bad value from a blocked rectangle.
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(2)/(SEQUENCE(2)-1)");

        book.Calculate();

        Assert.Equal(ExcelError.DivisionByZero, sheet["A1"].GetError());
        Assert.Equal(2d, sheet["A2"].GetNumber());
        Assert.Equal(RangeRef.Parse("A1:A2"), sheet.Cell("A1").SpillRange);
    }

    [Fact]
    public void SortAndFilterSpillTheirResults()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["A1"] = 3d;
        sheet["A2"] = 1d;
        sheet["A3"] = 2d;
        sheet.Cell("C1").SetFormula("SORT(A1:A3)");

        book.Calculate();

        Assert.Equal(1d, sheet["C1"].GetNumber());
        Assert.Equal(2d, sheet["C2"].GetNumber());
        Assert.Equal(3d, sheet["C3"].GetNumber());
    }

    [Fact]
    public void ASpillReferenceCoversTheWholeRectangle()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(4)");
        sheet.Cell("C1").SetFormula("SUM(A1#)");

        book.Calculate();

        Assert.Equal(10d, sheet["C1"].GetNumber());
    }

    [Fact]
    public void ASpillReferenceFollowsTheRectangleWhenItGrows()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["B1"] = 2d;
        sheet.Cell("A1").SetFormula("SEQUENCE(B1)");
        sheet.Cell("C1").SetFormula("SUM(A1#)");
        book.Calculate();
        Assert.Equal(3d, sheet["C1"].GetNumber());

        sheet["B1"] = 4d;
        book.Calculate();

        Assert.Equal(10d, sheet["C1"].GetNumber());
    }

    [Fact]
    public void ASpillReferenceToACellThatSpilledNothingIsRef()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["A1"] = 42d;
        sheet.Cell("C1").SetFormula("SUM(A1#)");

        book.Calculate();

        Assert.Equal(ExcelError.Reference, sheet["C1"].GetError());
    }

    [Fact]
    public void ASpillReferenceSurvivesARoundTripThroughText()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        sheet.Cell("C1").SetFormula("SUM(A1#)");

        book.Calculate();

        // The formula reads back with its suffix: the spelling is part of the meaning.
        Assert.Equal("SUM(A1#)", sheet.Cell("C1").Formula);
    }

    [Fact]
    public void ALegacyCseArrayStaysOneCell()
    {
        Worksheet sheet = Sheet(out Workbook book);
        sheet["A1"] = 1d;
        sheet["A2"] = 2d;

        // A CSE group read from a file: the master keeps its ref and does not spill.
        sheet.Cell("C1").SetFormula("SEQUENCE(3)");
        sheet.FormulaGroups.Add(new FormulaGroup(
            FormulaGroupKind.Array, CellRef.Parse("C1"), RangeRef.Parse("C1"), "SEQUENCE(3)"));

        book.Calculate();

        Assert.Equal(1d, sheet["C1"].GetNumber());
        Assert.True(sheet["C2"].IsBlank);
    }
}
