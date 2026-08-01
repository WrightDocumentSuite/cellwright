using System.IO.Compression;
using System.Text;
using Cellwright.Formulas;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Dynamic arrays through the file: a spilled formula is written as an array formula with the
/// metadata that says it is dynamic, and comes back a spill rather than a CSE relic.
/// </summary>
public sealed class SpillRoundTripTests
{
    private static async Task<Workbook> RoundTripAsync(Workbook book)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        return await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static async Task<string> PartOfAsync(Workbook book, string suffix)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (!entry.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                continue;

            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
        }

        return string.Empty;
    }

    [Fact]
    public async Task ASpillIsWrittenAsADynamicArrayFormula()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();

        string sheetXml = await PartOfAsync(book, "sheet1.xml");
        string metadata = await PartOfAsync(book, "metadata.xml");

        Assert.Contains("t=\"array\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("ref=\"A1:A3\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("cm=\"1\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("XLDAPR", metadata, StringComparison.Ordinal);
        Assert.Contains("dynamicArrayProperties", metadata, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ASpillComesBackASpillNotACseFormula()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();

        Workbook reloaded = await RoundTripAsync(book);
        Worksheet read = reloaded.Sheets[0];

        // The rectangle is known before any recalculation: the file said where it spilled.
        Assert.Equal(RangeRef.Parse("A1:A3"), read.Cell("A1").SpillRange);
        Assert.True(read.Cell("A2").IsSpilled);
        Assert.Equal(2d, read["A2"].GetNumber());
    }

    [Fact]
    public async Task AReloadedSpillRecalculatesOverItsOwnCells()
    {
        // The cached values occupy the rectangle; a recalculation must recognise them as the
        // formula's own output rather than somebody's cells standing in the way.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");
        book.Calculate();

        Workbook reloaded = await RoundTripAsync(book);
        reloaded.Calculate();

        Assert.Equal(3d, reloaded.Sheets[0]["A3"].GetNumber());
        Assert.NotEqual(ExcelError.Spill, reloaded.Sheets[0]["A1"].Kind == CellValueKind.Error
            ? reloaded.Sheets[0]["A1"].GetError()
            : default);
    }

    [Fact]
    public async Task ASpillReferenceSurvivesTheFile()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(4)");
        sheet.Cell("C1").SetFormula("SUM(A1#)");
        book.Calculate();

        Workbook reloaded = await RoundTripAsync(book);
        reloaded.Calculate();

        Assert.Equal("SUM(A1#)", reloaded.Sheets[0].Cell("C1").Formula);
        Assert.Equal(10d, reloaded.Sheets[0]["C1"].GetNumber());
    }

    [Fact]
    public async Task ACseArrayFormulaStaysCse()
    {
        // No metadata part, no cm: an array formula without them is the legacy kind and must
        // not start spilling on reload.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("C1").SetFormula("SEQUENCE(3)");
        sheet.FormulaGroups.Add(new FormulaGroup(
            FormulaGroupKind.Array, CellRef.Parse("C1"), RangeRef.Parse("C1"), "SEQUENCE(3)"));
        book.Calculate();

        Workbook reloaded = await RoundTripAsync(book);

        Assert.Null(reloaded.Sheets[0].Cell("C1").SpillRange);
        Assert.Single(reloaded.Sheets[0].FormulaGroups);
    }

    [Fact]
    public async Task AnUncalculatedWorkbookWritesPlainFormulasAndNoMetadata()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(3)");

        string sheetXml = await PartOfAsync(book, "sheet1.xml");
        string metadata = await PartOfAsync(book, "metadata.xml");

        // Nothing spilled, so there is nothing to mark; Excel recalculates on open and spills
        // for itself, which the full-calc flag already asks it to.
        Assert.DoesNotContain("t=\"array\"", sheetXml, StringComparison.Ordinal);
        Assert.Equal(string.Empty, metadata);
    }

    [Fact]
    public async Task TheMetadataPartIsNamedFromTheWorkbook()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetFormula("SEQUENCE(2)");
        book.Calculate();

        string relationships = await PartOfAsync(book, "workbook.xml.rels");

        Assert.Contains("sheetMetadata", relationships, StringComparison.Ordinal);
        Assert.Contains("metadata.xml", relationships, StringComparison.Ordinal);
    }
}
