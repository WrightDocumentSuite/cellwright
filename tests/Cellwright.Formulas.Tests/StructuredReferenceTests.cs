using Cellwright.Formulas.Parsing;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Structured references: <c>Table[Column]</c> and its relatives. They name a table rather than
/// coordinates, so what matters is that they follow the table and read the right band of it.
/// </summary>
public sealed class StructuredReferenceTests
{
    private static Worksheet Sheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = "Item";
        sheet["B1"] = "Qty";
        sheet["C1"] = "Price";
        for (int row = 2; row <= 5; row++)
        {
            sheet[row, 1] = $"item{row - 1}";
            sheet[row, 2] = row - 1;
            sheet[row, 3] = (row - 1) * 100;
        }

        sheet[6, 1] = "Total";
        sheet[6, 2] = 10;
        sheet[6, 3] = 1000;
        sheet.Tables.Add(new ExcelTable
        {
            Name = "Sales",
            Range = RangeRef.Parse("A1:C6"),
            Columns = ["Item", "Qty", "Price"],
            HasHeaderRow = true,
            HasTotalsRow = true,
        });

        return sheet;
    }

    [Theory]
    [InlineData("SUM(Sales[Qty])", 10.0)]
    [InlineData("SUM(Sales[Price])", 1000.0)]
    [InlineData("COUNT(Sales[Qty])", 4.0)]
    [InlineData("SUM(Sales[[#Data],[Qty]])", 10.0)]
    [InlineData("SUM(Sales[[#Totals],[Qty]])", 10.0)]
    [InlineData("SUM(Sales[[#All],[Qty]])", 20.0)]
    [InlineData("SUM(Sales[[Qty]:[Price]])", 1010.0)]
    [InlineData("COUNTA(Sales[[#Headers],[Qty]])", 1.0)]
    [InlineData("SUM(Sales[#Data])", 1010.0)]
    public void BandsAndColumnsResolveAgainstTheTable(string formula, double expected)
    {
        Assert.Equal(expected, Sheet().Evaluate(formula, "E1").GetNumber(), precision: 8);
    }

    [Fact]
    public void ThisRowReadsTheRowTheFormulaSitsOn()
    {
        Worksheet sheet = Sheet();

        // Naming the table works from anywhere; the bare form only inside the table itself.
        Assert.Equal(3.0, sheet.Evaluate("Sales[@Qty]", "E4").GetNumber());
        Assert.Equal(300.0, sheet.Evaluate("[@Price]", "A4").GetNumber());
        Assert.Equal(900.0, sheet.Evaluate("[@Qty]*[@Price]", "A4").GetNumber());
    }

    [Fact]
    public void ColumnNamesWithSpacesAndEscapesWork()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = "Unit price";
        sheet["A2"] = 25;
        sheet.Tables.Add(new ExcelTable
        {
            Name = "Items",
            Range = RangeRef.Parse("A1:A2"),
            Columns = ["Unit price"],
        });

        Assert.Equal(25.0, sheet.Evaluate("SUM(Items[Unit price])", "C1").GetNumber());
        Assert.Equal(25.0, sheet.Evaluate("SUM(Items[[Unit price]])", "C1").GetNumber());
    }

    [Fact]
    public void UnknownTablesAndColumnsGiveReferenceErrors()
    {
        Worksheet sheet = Sheet();

        Assert.Equal(ExcelError.Reference, sheet.Evaluate("SUM(Missing[Qty])", "E1").GetError());
        Assert.Equal(ExcelError.Reference, sheet.Evaluate("SUM(Sales[Nope])", "E1").GetError());

        // A bare this-row reference outside the table has no table to read.
        Assert.Equal(ExcelError.Reference, sheet.Evaluate("[@Qty]", "E20").GetError());
        Assert.Equal(ExcelError.Reference, sheet.Evaluate("Sales[@Qty]", "E20").GetError());
    }

    [Fact]
    public void ReferencesFollowTheTableWhenRowsAreInserted()
    {
        Worksheet sheet = Sheet();
        sheet.Cell("E1").SetFormula("SUM(Sales[Qty])");
        sheet.Workbook.Calculate();
        Assert.Equal(10.0, sheet["E1"].GetNumber());

        sheet.InsertRows(3);
        sheet["B3"] = 7;
        sheet.Workbook.Calculate();

        // The table grew a row and the formula, untouched, now covers it as well.
        Assert.Equal("SUM(Sales[Qty])", sheet.Cell("E1").Formula);
        Assert.Equal(17.0, sheet["E1"].GetNumber());
    }

    [Fact]
    public void EditingACellInsideTheTableRecalculatesWhatReadsIt()
    {
        Worksheet sheet = Sheet();
        sheet.Cell("E1").SetFormula("SUM(Sales[Qty])");
        sheet.Workbook.Calculate();

        sheet["B3"] = 20;
        sheet.Workbook.Calculate();

        Assert.Equal(28.0, sheet["E1"].GetNumber());
    }

    [Theory]
    [InlineData("SUM(Sales[Qty])")]
    [InlineData("Sales[@Qty]*2")]
    [InlineData("SUM(Sales[[#All],[Qty]:[Price]])")]
    [InlineData("[@[Unit price]]")]
    public void RenderingReproducesTheOriginalSpelling(string formula)
    {
        FormulaProgram program = FormulaParser.Parse(formula, Workbook.Create());

        Assert.Equal(formula, FormulaRenderer.Render(program, null));
    }

    [Fact]
    public void ParsingSplitsTheSpecIntoBandsAndColumns()
    {
        Assert.Equal(
            new TableRef("Sales", "Qty", null, TableArea.None, "Sales[Qty]"),
            TableRef.Parse("Sales", "[Qty]", "Sales[Qty]"));

        Assert.Equal(
            new TableRef("Sales", "Qty", "Price", TableArea.All, "Sales[[#All],[Qty]:[Price]]"),
            TableRef.Parse("Sales", "[[#All],[Qty]:[Price]]", "Sales[[#All],[Qty]:[Price]]"));

        Assert.Equal(
            new TableRef(null, "Unit price", null, TableArea.ThisRow, "[@[Unit price]]"),
            TableRef.Parse(null, "[@[Unit price]]", "[@[Unit price]]"));
    }
}
