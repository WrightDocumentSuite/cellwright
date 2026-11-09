using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// A loaded table or rule is written back from its original XML until something marks it edited.
/// Every public setter therefore has to do that marking, or a successful assignment is quietly
/// undone by the next save. Each case changes one thing on a reloaded workbook and reads it back.
/// </summary>
public sealed class LoadedEditPersistenceTests
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFUlEQVR4nGP8z8DwnwEJMKEL0FMQAG" +
        "0lAgcqA5wCAAAAAElFTkSuQmCC");

    private static async Task<MemoryStream> SaveAsync(Workbook workbook)
    {
        var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        return stream;
    }

    private static async Task<Workbook> LoadAsync(Stream stream)
    {
        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using MemoryStream stream = await SaveAsync(workbook);
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        return await LoadAsync(stream);
    }

    private static async Task<string> TableXmlAsync(Workbook workbook)
    {
        using MemoryStream stream = await SaveAsync(workbook);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.Entries.Single(e => e.FullName.StartsWith("xl/tables/", StringComparison.Ordinal));
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }

    private static Workbook TableBook(string range = "A1:B3")
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = "Name";
        sheet["B1"] = "Value";
        sheet["A2"] = "a";
        sheet["B2"] = 1;
        sheet["A3"] = "b";
        sheet["B3"] = 2;
        sheet["A4"] = "c";
        sheet["B4"] = 3;
        sheet.Tables.Add("Sales", range);
        return workbook;
    }

    [Fact]
    public async Task EditingARangeOfALoadedTableIsSaved()
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());

        loaded.Sheets["Data"].Tables.Find("Sales")!.Range = RangeRef.Parse("A1:B4");
        Workbook reloaded = await SaveAndReloadAsync(loaded);

        Assert.Equal(RangeRef.Parse("A1:B4"), reloaded.Sheets["Data"].Tables.Find("Sales")!.Range);
    }

    [Fact]
    public async Task SavingAnUntouchedTableKeepsItsOriginalPart()
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());

        Assert.False(loaded.Sheets["Data"].Tables.IsModified);
        Worksheet sheet = loaded.Sheets["Data"];
        ExcelTable table = sheet.Tables.Find("Sales")!;
        table.Range = table.Range;

        Assert.False(sheet.Tables.IsModified);
    }

    [Fact]
    public async Task ATableRangeCannotChangeWidthWithoutItsColumns()
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());
        ExcelTable table = loaded.Sheets["Data"].Tables.Find("Sales")!;

        Assert.Throws<ArgumentException>(() => table.Range = RangeRef.Parse("A1:C3"));
    }

    private static Workbook ValidationBook()
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = 1;
        sheet.Validations.Add(DataValidation.WholeBetween("A1:A3", 1, 10));
        return workbook;
    }

    [Fact]
    public async Task EditingFormulasOfALoadedValidationIsSaved()
    {
        Workbook loaded = await SaveAndReloadAsync(ValidationBook());

        loaded.Sheets["Data"].Validations[0].Formula2 = "100";
        Workbook reloaded = await SaveAndReloadAsync(loaded);

        Assert.Equal("100", reloaded.Sheets["Data"].Validations[0].Formula2);
    }

    [Fact]
    public async Task EditingTheRangesOfALoadedValidationIsSaved()
    {
        Workbook loaded = await SaveAndReloadAsync(ValidationBook());

        loaded.Sheets["Data"].Validations[0].Ranges = DataValidation.ParseRanges("C1:C5");
        Workbook reloaded = await SaveAndReloadAsync(loaded);

        Assert.Equal(RangeRef.Parse("C1:C5"), Assert.Single(reloaded.Sheets["Data"].Validations[0].Ranges));
    }

    [Fact]
    public async Task AssigningTheSameFormulaDoesNotCountAsAnEdit()
    {
        Workbook loaded = await SaveAndReloadAsync(ValidationBook());
        DataValidation validation = loaded.Sheets["Data"].Validations[0];

        validation.Formula1 = validation.Formula1;

        Assert.False(loaded.Sheets["Data"].Validations.IsModified);
    }

    [Fact]
    public async Task MovingALoadedPictureIsSaved()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Data").Pictures.Add(Png, "B2", "C3");
        Workbook loaded = await SaveAndReloadAsync(workbook);

        loaded.Sheets["Data"].Pictures[0].Anchor.From = CellRef.Parse("D4");
        loaded.Sheets["Data"].Pictures[0].Anchor.To = CellRef.Parse("E5");
        Workbook reloaded = await SaveAndReloadAsync(loaded);

        Assert.Equal(CellRef.Parse("D4"), reloaded.Sheets["Data"].Pictures[0].Anchor.From);
        Assert.Equal(CellRef.Parse("E5"), reloaded.Sheets["Data"].Pictures[0].Anchor.To);
    }

    [Fact]
    public async Task MovingALoadedNoteIsSavedAndFoundAtItsNewCell()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Data").Comments.Add("A1", "note");
        Workbook loaded = await SaveAndReloadAsync(workbook);

        CommentCollection notes = loaded.Sheets["Data"].Comments;
        notes[CellRef.Parse("A1")]!.Cell = CellRef.Parse("B2");

        Assert.Null(notes[CellRef.Parse("A1")]);
        Assert.Equal("note", notes[CellRef.Parse("B2")]!.Text);

        Workbook reloaded = await SaveAndReloadAsync(loaded);
        Assert.Equal("note", reloaded.Sheets["Data"].Comments[CellRef.Parse("B2")]!.Text);
        Assert.Null(reloaded.Sheets["Data"].Comments[CellRef.Parse("A1")]);
    }

    [Theory]
    [InlineData(1, "C1:D3", new[] { "Name", "Value" })]
    [InlineData(2, "A1:C3", new[] { "Name", "Column1", "Value" })]
    [InlineData(3, "A1:B3", new[] { "Name", "Value" })]
    public async Task InsertingAColumnKeepsTheTableConsistent(int column, string expectedRange, string[] expectedColumns)
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());
        Worksheet sheet = loaded.Sheets["Data"];

        sheet.InsertColumns(column, column == 1 ? 2 : 1);
        ExcelTable table = sheet.Tables.Find("Sales")!;

        Assert.Equal(RangeRef.Parse(expectedRange), table.Range);
        Assert.Equal(expectedColumns, table.Columns);

        for (int i = 0; i < table.Columns.Count; i++)
            Assert.Equal(table.Columns[i], sheet[table.Range.Start.Row, table.Range.Start.Column + i].ToString());

        Workbook reloaded = await SaveAndReloadAsync(loaded);
        ExcelTable after = reloaded.Sheets["Data"].Tables.Find("Sales")!;
        Assert.Equal(table.Range, after.Range);
        Assert.Equal(table.Columns, after.Columns);
    }

    [Fact]
    public async Task InsertingSeveralColumnsGivesEachAUniqueName()
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = "Column1";
        sheet["B1"] = "Value";
        sheet["A2"] = 1;
        sheet["B2"] = 2;
        sheet.Tables.Add("Sales", "A1:B2");

        sheet.InsertColumns(2, 2);

        ExcelTable table = sheet.Tables.Find("Sales")!;
        Assert.Equal(["Column1", "Column2", "Column3", "Value"], table.Columns);

        Workbook reloaded = await SaveAndReloadAsync(workbook);
        Assert.Equal(table.Columns, reloaded.Sheets["Data"].Tables.Find("Sales")!.Columns);
    }

    [Theory]
    [InlineData(1, "A1:A3", new[] { "Value" })]
    [InlineData(2, "A1:A3", new[] { "Name" })]
    public async Task DeletingAColumnKeepsTheTableConsistent(int column, string expectedRange, string[] expectedColumns)
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());
        Worksheet sheet = loaded.Sheets["Data"];

        sheet.DeleteColumns(column);
        ExcelTable table = sheet.Tables.Find("Sales")!;

        Assert.Equal(RangeRef.Parse(expectedRange), table.Range);
        Assert.Equal(expectedColumns, table.Columns);

        Workbook reloaded = await SaveAndReloadAsync(loaded);
        Assert.Equal(expectedColumns, reloaded.Sheets["Data"].Tables.Find("Sales")!.Columns);
    }

    [Fact]
    public async Task TheSavedTablePartAgreesWithTheRangeAfterAColumnInsert()
    {
        Workbook loaded = await SaveAndReloadAsync(TableBook());
        loaded.Sheets["Data"].InsertColumns(2);

        string xml = await TableXmlAsync(loaded);

        Assert.Contains("ref=\"A1:C3\"", xml, StringComparison.Ordinal);
        Assert.Contains("tableColumns count=\"3\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AColumnInsertKeepsATotalsRowInsideTheTable()
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = "Name";
        sheet["B1"] = "Value";
        sheet["A2"] = "a";
        sheet["B2"] = 1;
        sheet["A3"] = "Total";
        sheet.Tables.Add(new ExcelTable
        {
            Name = "Sales",
            Range = RangeRef.Parse("A1:B3"),
            Columns = ["Name", "Value"],
            HasTotalsRow = true,
        });

        sheet.InsertColumns(2);

        Workbook reloaded = await SaveAndReloadAsync(workbook);
        ExcelTable table = reloaded.Sheets["Data"].Tables.Find("Sales")!;
        Assert.Equal(RangeRef.Parse("A1:C3"), table.Range);
        Assert.Equal(3, table.Columns.Count);
        Assert.True(table.HasTotalsRow);
    }
}
