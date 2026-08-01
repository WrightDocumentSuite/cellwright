using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// What a chart costs in the binary format. There is no preserving it: a chart's records point
/// at fonts, formats and strings by their position in tables this library rebuilds, so copying
/// them across would leave a chart plotting whatever now sits at those positions. They are
/// dropped instead — the point of these tests is that nothing else goes with them.
/// </summary>
public sealed class ChartSheetTests
{
    private static string? Find(string name) =>
        XlsCorpus.Files.FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    [Fact]
    public async Task AChartSheetIsSkippedAndSaidToBeSkipped()
    {
        string? path = Find("44010-SingleChart.xls");
        Assert.SkipWhen(path is null, "44010-SingleChart.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var warnings = new List<WorkbookWarning>();
        Workbook workbook = await XlsWorkbook.LoadAsync(
            path!, new XlsLoadOptions { OnWarning = warnings.Add }, TestContext.Current.CancellationToken);

        Assert.NotEmpty(workbook.Sheets);
        Assert.Contains(warnings, static warning => warning.Code == "SkippedSheet");
        Assert.All(workbook.Sheets, static sheet => Assert.NotEqual(0, sheet.Name.Length));
    }

    /// <summary>
    /// A chart sheet in the middle of the list shifts every sheet after it, and a name scoped
    /// to one of those would end up on the wrong sheet if the two orders were confused.
    /// </summary>
    [Fact]
    public async Task SheetsAfterAChartSheetKeepTheirContentAndTheirNames()
    {
        string? path = Find("44010-TwoCharts.xls");
        Assert.SkipWhen(path is null, "44010-TwoCharts.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);

        foreach (DefinedName name in workbook.Names)
        {
            if (name.Scope is { } scope)
                Assert.Same(scope, workbook.Sheets.Find(scope.Name));
        }

        // Whatever survived has to be a grid this library can save again.
        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        Workbook reloaded = await XlsWorkbook.LoadAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal(workbook.Sheets.Count, reloaded.Sheets.Count);
    }

    /// <summary>
    /// An embedded chart lives among the drawing records of an ordinary sheet, mixed in with
    /// the shapes that carry notes. Its objects must not be mistaken for notes.
    /// </summary>
    [Fact]
    public async Task AnEmbeddedChartDoesNotTurnIntoANote()
    {
        string? path = Find("EmbeddedChartHeaderTest.xls");
        Assert.SkipWhen(path is null, "EmbeddedChartHeaderTest.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);

        foreach (Worksheet sheet in workbook.Sheets)
        {
            foreach (CellComment comment in sheet.Comments)
                Assert.False(string.IsNullOrEmpty(comment.Text), $"A shape on '{sheet.Name}' became an empty note.");
        }
    }
}
