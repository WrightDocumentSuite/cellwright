using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Shared formulas are how Excel stores a filled-down column: the first cell holds the text and
/// the others only say which group they are in. Saving has to keep what they compute whether or
/// not anything ever calculated the workbook.
/// </summary>
public sealed class SharedFormulaPersistenceTests
{
    private const string SharedColumn =
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>2</v></c></row>" +
        "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>" +
        "<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row>";

    /// <summary>The range encloses B2, which holds a plain value and belongs to no group.</summary>
    private const string SharedColumnWithAHole =
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>2</v></c></row>" +
        "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><v>99</v></c></row>" +
        "<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row>";

    private static async Task<MemoryStream> PackageWithRowsAsync(string rows)
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("S")["A1"] = 1;

        var saved = new MemoryStream();
        await workbook.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using (var archive = new ZipArchive(saved, ZipArchiveMode.Update, leaveOpen: true))
        {
            ZipArchiveEntry sheetEntry = archive.Entries.Single(static e => e.FullName.Contains("worksheets/sheet", StringComparison.Ordinal));
            string xml;
            using (var reader = new StreamReader(sheetEntry.Open(), Encoding.UTF8))
                xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);

            int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
            int end = xml.IndexOf("</sheetData>", StringComparison.Ordinal);
            string rewritten = string.Concat(xml.AsSpan(0, start), $"<sheetData>{rows}", xml.AsSpan(end));

            sheetEntry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(sheetEntry.FullName).Open(), new UTF8Encoding(false));
            await writer.WriteAsync(rewritten);
        }

        saved.Position = 0;
        return saved;
    }

    private static async Task<Workbook> LoadAsync(Stream stream)
    {
        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        return await LoadAsync(stream);
    }

    [Fact]
    public async Task SavingWithoutCalculatingKeepsTheFormulasOfMembers()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumn);

        Workbook reloaded = await SaveAndReloadAsync(await LoadAsync(package));
        Worksheet sheet = reloaded.Sheets["S"];

        Assert.Equal("A1*2", sheet.Cell("B1").Formula);
        Assert.Equal("A2*2", sheet.Cell("B2").Formula);
        Assert.Equal("A3*2", sheet.Cell("B3").Formula);
        Assert.Equal(4.0, sheet["B2"].GetNumber());
    }

    [Fact]
    public async Task FormulasKeptBySavingStillCalculate()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumn);
        Workbook reloaded = await SaveAndReloadAsync(await LoadAsync(package));
        Worksheet sheet = reloaded.Sheets["S"];

        sheet["A3"] = 10;
        reloaded.Calculate();

        Assert.Equal(20.0, sheet["B3"].GetNumber());
    }

    [Fact]
    public async Task SavingTwiceWithoutCalculatingChangesNothing()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumn);
        Workbook loaded = await LoadAsync(package);

        Workbook once = await SaveAndReloadAsync(loaded);
        Workbook twice = await SaveAndReloadAsync(once);

        Assert.Equal("A3*2", twice.Sheets["S"].Cell("B3").Formula);
    }

    [Fact]
    public async Task SavingLeavesAPlainValueInsideTheRangeAlone()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumnWithAHole);

        Workbook reloaded = await SaveAndReloadAsync(await LoadAsync(package));
        Worksheet sheet = reloaded.Sheets["S"];

        Assert.Null(sheet.Cell("B2").Formula);
        Assert.Equal(99.0, sheet["B2"].GetNumber());
        Assert.Equal("A3*2", sheet.Cell("B3").Formula);
    }

    [Fact]
    public async Task CalculatingLeavesAPlainValueInsideTheRangeAlone()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumnWithAHole);
        Workbook loaded = await LoadAsync(package);

        loaded.Calculate();
        Worksheet sheet = loaded.Sheets["S"];

        Assert.Null(sheet.Cell("B2").Formula);
        Assert.Equal(99.0, sheet["B2"].GetNumber());
        Assert.Equal(6.0, sheet["B3"].GetNumber());
    }

    [Fact]
    public async Task CopyingASheetBeforeSavingKeepsMembers()
    {
        using MemoryStream package = await PackageWithRowsAsync(SharedColumn);
        Workbook loaded = await LoadAsync(package);

        loaded.Sheets.Copy("S", "Copy");
        Workbook reloaded = await SaveAndReloadAsync(loaded);

        Assert.Equal("A3*2", reloaded.Sheets["Copy"].Cell("B3").Formula);
    }
}
