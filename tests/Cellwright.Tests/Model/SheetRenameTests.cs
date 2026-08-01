using Xunit;
using Cellwright.IO;

namespace Cellwright.Tests.Model;

/// <summary>
/// Renaming a sheet has to leave the workbook working: every formula, name, rule and link that
/// spelled the old name out must follow it.
/// </summary>
public sealed class SheetRenameTests
{
    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet data = book.Sheets.Add("Data");
        Worksheet report = book.Sheets.Add("Report");
        data["A1"] = 10;
        data["A2"] = 20;
        report.Cell("B1").SetFormula("SUM(Data!A1:A2)");
        report.Cell("B2").SetFormula("Data!A1+A2");
        return book;
    }

    [Fact]
    public void FormulasOfEverySheetFollowTheNewName()
    {
        Workbook book = Sample();

        book.Sheets["Data"].Name = "Facts";

        Worksheet report = book.Sheets["Report"];
        Assert.Equal("SUM(Facts!A1:A2)", report.Cell("B1").Formula);
        Assert.Equal("Facts!A1+A2", report.Cell("B2").Formula);
    }

    [Fact]
    public void NameNeedingQuotesGetsThem()
    {
        Workbook book = Sample();

        book.Sheets["Data"].Name = "Q1 2026";

        Assert.Equal("SUM('Q1 2026'!A1:A2)", book.Sheets["Report"].Cell("B1").Formula);
    }

    [Fact]
    public void DefinedNamesFollowTheNewName()
    {
        Workbook book = Sample();
        book.Names.Add("Rate", "Data!$A$1");
        book.Sheets["Data"].SetPrintArea("A1:B9");

        book.Sheets["Data"].Name = "Facts";

        Assert.Equal("Facts!$A$1", book.Names.Find("Rate")?.RefersTo);
        Assert.Equal(RangeRef.Parse("A1:B9"), book.Sheets["Facts"].PrintArea);
    }

    [Fact]
    public void ValidationsConditionalFormatsAndLinksFollowTheNewName()
    {
        Workbook book = Sample();
        Worksheet report = book.Sheets["Report"];
        report.Validations.Add(DataValidation.List("C1:C9", "Data!$A$1:$A$2"));
        report.ConditionalFormats.Add(ConditionalFormatRule.Expression("D1:D9", "D1>Data!$A$1", DifferentialStyle.Bad));
        report.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("E1"), "Data!A1"));

        book.Sheets["Data"].Name = "Facts";

        Assert.Equal("Facts!$A$1:$A$2", report.Validations[0].Formula1);
        Assert.Equal("D1>Facts!$A$1", report.ConditionalFormats[0].Formulas[0]);
        Assert.Equal("Facts!A1", report.Hyperlinks[0].Location);
    }

    [Fact]
    public void RenamingToTheSameNameChangesNothing()
    {
        Workbook book = Sample();

        book.Sheets["Data"].Name = "Data";

        Assert.Equal("SUM(Data!A1:A2)", book.Sheets["Report"].Cell("B1").Formula);
    }

    [Fact]
    public void ASheetNamedInsideAStringIsNotTouched()
    {
        Workbook book = Sample();
        book.Sheets["Report"].Cell("B3").SetFormula("\"Data!A1\"&Data!A1");

        book.Sheets["Data"].Name = "Facts";

        Assert.Equal("\"Data!A1\"&Facts!A1", book.Sheets["Report"].Cell("B3").Formula);
    }

    [Fact]
    public async Task RenamingRetargetsTheChartsOfALoadedWorkbook()
    {
        string? path = SampleCorpus.Files(int.MaxValue)
            .FirstOrDefault(static file => file.EndsWith("ChartsheetAndPivotTable.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "The sample workbook with a chart was not found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        Workbook book = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        string original = book.Sheets[0].Name;
        book.Sheets[0].Name = "Renamed";

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        string charts = ReadChartParts(saved);
        Assert.DoesNotContain($"<c:f>{original}!", charts, StringComparison.Ordinal);
        Assert.Contains("<c:f>Renamed!", charts, StringComparison.Ordinal);
    }

    [Fact]
    public void ForeignXmlIsRetargetedByNamespaceRatherThanByItsChosenPrefix()
    {
        var package = new PreservedPackage
        {
            ContentTypes = new ContentTypeMap(),
            WorkbookPartPath = XlsxSchema.PartWorkbook,
            WorkbookContentType = XlsxSchema.ContentTypeWorkbook,
        };
        package.Parts["/xl/charts/chart1.xml"] = System.Text.Encoding.UTF8.GetBytes("""
            <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <ser><f>Data!$A$1:$A$2</f></ser>
            </chartSpace>
            """);
        package.Parts["/xl/pivotCache/pivotCacheDefinition1.xml"] = System.Text.Encoding.UTF8.GetBytes("""
            <p:pivotCacheDefinition xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <p:cacheSource><p:worksheetSource
                  sheet='Data' ref='A1:B2'/></p:cacheSource>
            </p:pivotCacheDefinition>
            """);

        package.RetargetSheetName("Data", "Facts");

        string chart = System.Text.Encoding.UTF8.GetString(package.Parts["/xl/charts/chart1.xml"]);
        string pivot = System.Text.Encoding.UTF8.GetString(package.Parts["/xl/pivotCache/pivotCacheDefinition1.xml"]);
        Assert.Contains("Facts!$A$1:$A$2", chart, StringComparison.Ordinal);
        Assert.Contains("sheet=\"Facts\"", pivot, StringComparison.Ordinal);
    }

    private static string ReadChartParts(Stream stream)
    {
        using var archive = new System.IO.Compression.ZipArchive(
            stream, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        var text = new System.Text.StringBuilder();
        foreach (System.IO.Compression.ZipArchiveEntry entry in archive.Entries)
        {
            if (!entry.FullName.Contains("/charts/", StringComparison.OrdinalIgnoreCase))
                continue;

            using var reader = new StreamReader(entry.Open());
            text.Append(reader.ReadToEnd());
        }

        return text.ToString();
    }
}
