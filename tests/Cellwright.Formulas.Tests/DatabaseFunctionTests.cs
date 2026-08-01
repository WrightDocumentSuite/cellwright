using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The database functions, on the orchard table Excel's own documentation uses for them, so
/// every expected value here is one Microsoft publishes for the same arguments.
/// </summary>
public sealed class DatabaseFunctionTests
{
    /// <summary>
    /// Conditions in rows 1 to 3, the records themselves in rows 4 to 10. The two tables sit on
    /// one sheet because that is how these functions are always written.
    /// </summary>
    private static Worksheet Orchard()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");

        sheet["A1"] = "Tree";
        sheet["B1"] = "Height";
        sheet["C1"] = "Age";
        sheet["D1"] = "Yield";
        sheet["E1"] = "Profit";
        sheet["F1"] = "Height";
        sheet["A2"] = "Apple";
        sheet["B2"] = ">10";
        sheet["F2"] = "<16";
        sheet["A3"] = "Pear";

        sheet["A4"] = "Tree";
        sheet["B4"] = "Height";
        sheet["C4"] = "Age";
        sheet["D4"] = "Yield";
        sheet["E4"] = "Profit";

        Record(sheet, 5, "Apple", 18, 20, 14, 105);
        Record(sheet, 6, "Pear", 12, 12, 10, 96);
        Record(sheet, 7, "Cherry", 13, 14, 9, 105);
        Record(sheet, 8, "Apple", 14, 15, 10, 75);
        Record(sheet, 9, "Pear", 9, 8, 8, 76.80);
        Record(sheet, 10, "Apple", 8, 9, 6, 45);
        return sheet;
    }

    private static void Record(Worksheet sheet, int row, string tree, double height, double age, double yield, double profit)
    {
        sheet[row, 1] = CellValue.Text(tree);
        sheet[row, 2] = CellValue.Number(height);
        sheet[row, 3] = CellValue.Number(age);
        sheet[row, 4] = CellValue.Number(yield);
        sheet[row, 5] = CellValue.Number(profit);
    }

    private static CellValue Eval(string formula) => Orchard().Evaluate(formula, "H1");

    [Theory]
    [InlineData("DCOUNT(A4:E10,\"Age\",A1:F2)", 1.0)]
    [InlineData("DCOUNTA(A4:E10,\"Profit\",A1:F2)", 1.0)]
    [InlineData("DMAX(A4:E10,\"Profit\",A1:B2)", 105.0)]
    [InlineData("DMIN(A4:E10,\"Profit\",A1:B2)", 75.0)]
    [InlineData("DSUM(A4:E10,\"Profit\",A1:A2)", 225.0)]
    [InlineData("DSUM(A4:E10,\"Profit\",A1:F2)", 75.0)]
    [InlineData("DPRODUCT(A4:E10,\"Yield\",A1:B2)", 140.0)]
    [InlineData("DAVERAGE(A4:E10,\"Yield\",A1:B2)", 12.0)]
    [InlineData("DSTDEV(A4:E10,\"Yield\",A1:A3)", 2.96647939)]
    [InlineData("DSTDEVP(A4:E10,\"Yield\",A1:A3)", 2.65329983)]
    [InlineData("DVAR(A4:E10,\"Yield\",A1:A3)", 8.8)]
    [InlineData("DVARP(A4:E10,\"Yield\",A1:A3)", 7.04)]
    public void TheOrchardAnswersAsExcelDocumentsIt(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).GetNumber(), precision: 8);
    }

    /// <summary>
    /// Conditions on one row must all hold; conditions on separate rows are alternatives. Naming
    /// the same column twice is how a range is stated, since one row cannot hold two conditions
    /// under one heading.
    /// </summary>
    [Fact]
    public void ConditionsCombineAcrossColumnsAndAlternateDownRows()
    {
        // Apple, over ten feet and under sixteen: one tree of the three.
        Assert.Equal(1.0, Eval("DCOUNT(A4:E10,\"Age\",A1:F2)").GetNumber());

        // Apple or pear, whatever the height: five of the six.
        Assert.Equal(5.0, Eval("DCOUNT(A4:E10,\"Age\",A1:A3)").GetNumber());

        // Nothing but a heading states no condition, so every record is in.
        Assert.Equal(6.0, Eval("DCOUNT(A4:E10,\"Age\",A1:A1)").GetNumber());
    }

    /// <summary>A field is named or numbered, and the position counts from one.</summary>
    [Fact]
    public void TheFieldIsNamedOrNumbered()
    {
        Assert.Equal(
            Eval("DAVERAGE(A4:E10,\"Age\",A1:A2)").GetNumber(),
            Eval("DAVERAGE(A4:E10,3,A1:A2)").GetNumber());

        // Using the records as their own conditions admits all of them, which is the roundabout
        // way Excel's own documentation states "no filter at all".
        Assert.Equal(13.0, Eval("DAVERAGE(A4:E10,3,A4:E10)").GetNumber());

        Assert.Equal(ExcelError.Value, Eval("DSUM(A4:E10,\"Colour\",A1:A2)").GetError());
        Assert.Equal(ExcelError.Value, Eval("DSUM(A4:E10,9,A1:A2)").GetError());
    }

    /// <summary>
    /// <c>DGET</c> asks for the one record that fits, and tells apart the two ways that can fail.
    /// </summary>
    [Fact]
    public void GettingOneRecordDistinguishesNoneFromSeveral()
    {
        Assert.Equal(10.0, Eval("DGET(A4:E10,\"Yield\",A1:F2)").GetNumber());
        Assert.Equal(ExcelError.Number, Eval("DGET(A4:E10,\"Yield\",A1:A3)").GetError());
    }

    /// <summary>Counting records rather than values is the one case that needs no field.</summary>
    [Fact]
    public void AnOmittedFieldCountsTheRecordsThemselves()
    {
        Assert.Equal(5.0, Eval("DCOUNTA(A4:E10,,A1:A3)").GetNumber());
        Assert.Equal(1.0, Eval("DCOUNT(A4:E10,,A1:F2)").GetNumber());

        // There being no zeroth column, a field of zero says the same as leaving it out.
        Assert.Equal(5.0, Eval("DCOUNT(A4:E10,0,A1:A3)").GetNumber());
        Assert.Equal(5.0, Eval("DCOUNTA(A4:E10,0,A1:A3)").GetNumber());
    }

    /// <summary>Conditions nothing fits leave the aggregates with nothing to add up.</summary>
    [Fact]
    public void ConditionsNothingFitsLeaveTheAggregatesEmpty()
    {
        Worksheet sheet = Orchard();
        sheet["H5"] = "Tree";
        sheet["H6"] = "Plum";

        Assert.Equal(0.0, sheet.Evaluate("DSUM(A4:E10,\"Profit\",H5:H6)", "J1").GetNumber());
        Assert.Equal(0.0, sheet.Evaluate("DCOUNT(A4:E10,\"Profit\",H5:H6)", "J1").GetNumber());
        Assert.Equal(ExcelError.DivisionByZero, sheet.Evaluate("DAVERAGE(A4:E10,\"Profit\",H5:H6)", "J1").GetError());
        Assert.Equal(ExcelError.Value, sheet.Evaluate("DGET(A4:E10,\"Profit\",H5:H6)", "J1").GetError());
    }
}
