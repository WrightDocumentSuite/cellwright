using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// What the file says about itself. None of it touches a cell, which is why a workbook used to
/// be saved without any of it and nobody noticed until the file reached somebody who sorts by
/// author.
/// </summary>
public sealed class DocumentPropertiesTests
{
    private static async Task<MemoryStream> SaveAsync(Workbook workbook)
    {
        var stream = new MemoryStream();
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
        return stream;
    }

    private static async Task<Workbook> RoundTripAsync(Workbook workbook)
    {
        using MemoryStream saved = await SaveAsync(workbook);
        return await Workbook.LoadAsync(saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task ANewWorkbookCarriesTheDocumentPropertyPartsExcelWrites()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Data")["A1"] = "value";

        using MemoryStream saved = await SaveAsync(workbook);
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read);
        Assert.Contains(archive.Entries, static entry => entry.FullName == "docProps/core.xml");
        Assert.Contains(archive.Entries, static entry => entry.FullName == "docProps/app.xml");
        Assert.DoesNotContain(archive.Entries, static entry => entry.FullName == "docProps/custom.xml");
    }

    [Fact]
    public async Task TheFieldsSurviveASaveAndReload()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Data");
        workbook.Properties.Title = "Quarterly report";
        workbook.Properties.Subject = "Revenue";
        workbook.Properties.Author = "Анна Смирнова";
        workbook.Properties.Keywords = "revenue;q1";
        workbook.Properties.Description = "Numbers & commentary <draft>";
        workbook.Properties.Category = "Finance";
        workbook.Properties.ContentStatus = "Draft";
        workbook.Properties.LastModifiedBy = "Reviewer";
        workbook.Properties.Company = "Contoso";
        workbook.Properties.Manager = "Head of Finance";
        workbook.Properties.Created = new DateTime(2020, 3, 4, 5, 6, 7, DateTimeKind.Utc);

        WorkbookProperties reloaded = (await RoundTripAsync(workbook)).Properties;
        Assert.Equal("Quarterly report", reloaded.Title);
        Assert.Equal("Revenue", reloaded.Subject);
        Assert.Equal("Анна Смирнова", reloaded.Author);
        Assert.Equal("revenue;q1", reloaded.Keywords);
        Assert.Equal("Numbers & commentary <draft>", reloaded.Description);
        Assert.Equal("Finance", reloaded.Category);
        Assert.Equal("Draft", reloaded.ContentStatus);
        Assert.Equal("Reviewer", reloaded.LastModifiedBy);
        Assert.Equal("Contoso", reloaded.Company);
        Assert.Equal("Head of Finance", reloaded.Manager);
        Assert.Equal(new DateTime(2020, 3, 4, 5, 6, 7, DateTimeKind.Utc), reloaded.Created);
    }

    [Fact]
    public async Task CustomPropertiesKeepTheirTypes()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Data");
        workbook.Properties.Custom["Department"] = CellValue.Text("Finance");
        workbook.Properties.Custom["Revision"] = CellValue.Number(7);
        workbook.Properties.Custom["Approved"] = CellValue.Boolean(true);
        workbook.Properties.Custom["Signed"] = CellValue.DateTime(new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc));

        CustomPropertyCollection reloaded = (await RoundTripAsync(workbook)).Properties.Custom;
        Assert.Equal(4, reloaded.Count);
        Assert.Equal("Finance", reloaded["Department"].GetText());
        Assert.Equal(7, reloaded["Revision"].GetNumber());
        Assert.True(reloaded["Approved"].GetBoolean());
        Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc), reloaded["Signed"].GetDateTime());
    }

    [Fact]
    public async Task TheModifiedStampIsRefreshedOnSaveUnlessItWasChosen()
    {
        Workbook refreshed = Workbook.Create();
        refreshed.Sheets.Add("Data");
        DateTime before = DateTime.UtcNow.AddSeconds(-1);

        Assert.InRange((await RoundTripAsync(refreshed)).Properties.Modified!.Value, before, DateTime.UtcNow.AddMinutes(1));

        Workbook pinned = Workbook.Create();
        pinned.Sheets.Add("Data");
        pinned.Properties.Modified = new DateTime(2001, 2, 3, 4, 5, 6, DateTimeKind.Utc);

        Assert.Equal(
            new DateTime(2001, 2, 3, 4, 5, 6, DateTimeKind.Utc),
            (await RoundTripAsync(pinned)).Properties.Modified);
    }

    /// <summary>
    /// The extended part carries statistics no spreadsheet model has a place for. They are
    /// replayed rather than dropped — but the two lists that name the sheets are rebuilt,
    /// because the model is free to have renamed one since.
    /// </summary>
    [Fact]
    public async Task LoadedPropertiesKeepWhatTheModelDoesNotNameAndRefreshTheSheetList()
    {
        string? path = SampleCorpus.Files(int.MaxValue).FirstOrDefault(static file =>
            file.EndsWith("ChartsheetAndPivotTable.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "The sample workbook was not found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        Workbook workbook = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        workbook.Sheets[0].Name = "Renamed";

        // Saved without the validator: the chart this file carries uses an element the SDK's
        // schema predates, and that has nothing to do with what is being checked here.
        using var saved = new MemoryStream();
        await workbook.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        string app = ReadEntry(saved, "docProps/app.xml");

        Assert.Contains("<vt:lpstr>Renamed</vt:lpstr>", app, StringComparison.Ordinal);
        Assert.Contains("DocSecurity", app, StringComparison.Ordinal);
    }

    private static string ReadEntry(Stream package, string entryName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(archive.GetEntry(entryName)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
