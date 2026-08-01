using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Inserting and deleting rows and columns. What makes this hard is not the cells but
/// everything anchored to them, so each case checks one class of anchor as well as the values.
/// </summary>
public sealed class StructuralEditTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        for (int row = 1; row <= 5; row++)
        {
            sheet[row, 1] = $"row{row}";
            sheet[row, 2] = row * 10;
        }

        sheet.Cell("C1").SetFormula("SUM(B1:B5)");
        return book;
    }

    [Fact]
    public void InsertingRowsMovesCellsAndStretchesFormulas()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];

        sheet.InsertRows(2, 2);

        Assert.Equal("row1", sheet["A1"].GetText());
        Assert.True(sheet["A2"].IsBlank);
        Assert.Equal("row2", sheet["A4"].GetText());
        Assert.Equal(50.0, sheet["B7"].GetNumber());
        Assert.Equal("SUM(B1:B7)", sheet.Cell("C1").Formula);
    }

    [Fact]
    public void DeletingRowsPullsCellsUpAndShrinksFormulas()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];

        sheet.DeleteRows(2, 2);

        Assert.Equal("row1", sheet["A1"].GetText());
        Assert.Equal("row4", sheet["A2"].GetText());
        Assert.True(sheet["A4"].IsBlank);
        Assert.Equal("SUM(B1:B3)", sheet.Cell("C1").Formula);
    }

    [Fact]
    public void FormulaPointingIntoDeletedCellsBreaks()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Cell("D1").SetFormula("B3*2");

        sheet.DeleteRows(3);

        Assert.Equal("#REF!*2", sheet.Cell("D1").Formula);
    }

    [Fact]
    public void ColumnsMoveTheOtherAxis()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];

        sheet.InsertColumns(2);

        Assert.Equal("row1", sheet["A1"].GetText());
        Assert.True(sheet["B1"].IsBlank);
        Assert.Equal(10.0, sheet["C1"].GetNumber());
        Assert.Equal("SUM(C1:C5)", sheet.Cell("D1").Formula);
    }

    [Fact]
    public void FormulasOnOtherSheetsFollowTheEdit()
    {
        Workbook book = Sample();
        Worksheet report = book.Sheets.Add("Report");
        report.Cell("A1").SetFormula("Data!B3");
        report.Cell("A2").SetFormula("SUM(Data!B1:B5)");
        report.Cell("A3").SetFormula("B3");

        book.Sheets["Data"].InsertRows(2);

        Assert.Equal("Data!B4", report.Cell("A1").Formula);
        Assert.Equal("SUM(Data!B1:B6)", report.Cell("A2").Formula);
        Assert.Equal("B3", report.Cell("A3").Formula);
    }

    [Fact]
    public void DefinedNamesFollowTheEdit()
    {
        Workbook book = Sample();
        book.Names.Add("Amounts", "Data!$B$1:$B$5");
        book.Sheets[0].SetPrintArea("A1:C5");

        book.Sheets[0].InsertRows(1, 3);

        Assert.Equal("Data!$B$4:$B$8", book.Names.Find("Amounts")?.RefersTo);
        Assert.Equal(RangeRef.Parse("A4:C8"), book.Sheets[0].PrintArea);
    }

    [Fact]
    public void RowPropertiesMoveAndInsertedRowsTakeTheSizeOfTheOneAbove()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Row(2).SetHeight(30).Hide();
        sheet.Row(4).Hide();

        sheet.InsertRows(3, 2);

        Assert.Equal(30, sheet.Row(2).Height);
        Assert.True(sheet.Row(2).Hidden);
        Assert.True(sheet.Row(6).Hidden);

        // The new rows copy the height of row 2 but stay visible.
        Assert.Equal(30, sheet.Row(3).Height);
        Assert.False(sheet.Row(3).Hidden);
    }

    [Fact]
    public void MergesLinksAndNotesMoveWithTheirCells()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Merge(RangeRef.Parse("A3:B3"));
        sheet.Hyperlinks.Add(RangeRef.Parse("A4"), "https://example.com");
        sheet.Comments.Add("A5", "check me");

        sheet.InsertRows(2);

        Assert.Equal(RangeRef.Parse("A4:B4"), Assert.Single(sheet.MergedRanges));
        Assert.Equal(RangeRef.Parse("A5"), sheet.Hyperlinks[0].Range);
        Assert.Equal("check me", sheet.Comments[CellRef.Parse("A6")]?.Text);
    }

    [Fact]
    public void AnchorsInsideDeletedRowsGoAway()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Hyperlinks.Add(RangeRef.Parse("A3"), "https://example.com");
        sheet.Comments.Add("A3", "gone");
        sheet.Merge(RangeRef.Parse("A3:B3"));

        sheet.DeleteRows(3);

        Assert.Empty(sheet.Hyperlinks);
        Assert.Empty(sheet.Comments);
        Assert.Empty(sheet.MergedRanges);
    }

    [Fact]
    public void RulesAndTablesFollowTheEdit()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Validations.Add(DataValidation.WholeBetween("B1:B5", 1, 99));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression("A1:B5", "$B1>20", DifferentialStyle.Bad));
        sheet.Tables.Add("Rows", "A1:B5");
        sheet.AutoFilter("A1:B5");

        sheet.InsertRows(2);

        Assert.Equal(RangeRef.Parse("B1:B6"), Assert.Single(sheet.Validations[0].Ranges));
        Assert.Equal(RangeRef.Parse("A1:B6"), Assert.Single(sheet.ConditionalFormats[0].Ranges));
        Assert.Equal("$B1>20", sheet.ConditionalFormats[0].Formulas[0]);
        Assert.Equal(RangeRef.Parse("A1:B6"), sheet.Tables[0].Range);
        Assert.Equal(RangeRef.Parse("A1:B6"), sheet.AutoFilterRange);
    }

    [Fact]
    public void PageBreaksFollowTheEdit()
    {
        Workbook book = Sample();
        book.Sheets[0].PageSetup.AddRowBreak(4);
        book.Sheets[0].PageSetup.AddRowBreak(10);

        book.Sheets[0].InsertRows(2, 3);

        Assert.Equal([7, 13], book.Sheets[0].PageSetup.RowBreaks);
    }

    [Fact]
    public async Task TheEditedSheetIsStillAValidFileClosedXmlCanRead()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.Merge(RangeRef.Parse("A3:B3"));
        sheet.Comments.Add("A4", "note");
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B1:B5", 20, DifferentialStyle.Good));
        sheet.InsertRows(2, 2);
        sheet.DeleteColumns(1);

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(10.0, loaded["A1"].GetNumber());
        Assert.Equal("SUM(A1:A7)", loaded.Cell("B1").Formula);
        Assert.Equal(RangeRef.Parse("A1:A7"), Assert.Single(loaded.ConditionalFormats[0].Ranges));
    }

    [Fact]
    public void EditingWithNothingToMoveIsANoOp()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];

        sheet.InsertRows(3, 0);

        Assert.Equal("row3", sheet["A3"].GetText());
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.InsertRows(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.DeleteRows(1, -1));
    }
}
