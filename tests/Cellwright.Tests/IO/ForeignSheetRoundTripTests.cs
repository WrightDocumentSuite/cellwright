using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>Workbook sheet positions include chart sheets even though the model does not.</summary>
public sealed class ForeignSheetRoundTripTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Fact]
    public async Task ALocalNameKeepsItsFullSheetPositionAfterAChartSheet()
    {
        using MemoryStream source = Package(includeWorksheet: true);
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        DefinedName name = Assert.Single(book.Names);
        Assert.Same(book.Sheets["Data"], name.Scope);

        using var saved = new MemoryStream();
        await book.SaveAsync(
            saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        XDocument workbookXml = ReadWorkbook(saved);
        XElement written = Assert.Single(workbookXml.Descendants(Main + "definedName"));
        Assert.Equal("1", written.Attribute("localSheetId")?.Value);

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Same(reloaded.Sheets["Data"], Assert.Single(reloaded.Names).Scope);
    }

    [Fact]
    public async Task AWorkbookContainingOnlyAChartSheetDoesNotGainAnEmptyWorksheet()
    {
        using MemoryStream source = Package(includeWorksheet: false);
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Empty(book.Sheets);

        using var saved = new MemoryStream();
        await book.SaveAsync(
            saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        XDocument workbookXml = ReadWorkbook(saved);
        XElement sheet = Assert.Single(workbookXml.Descendants(Main + "sheet"));
        Assert.Equal("Chart", sheet.Attribute("name")?.Value);
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.Contains("worksheets/", StringComparison.Ordinal));
    }

    private static XDocument ReadWorkbook(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.xml")).Open();
        return XDocument.Load(xml);
    }

    private static MemoryStream Package(bool includeWorksheet)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            string worksheetType = includeWorksheet
                ? "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                : "";
            Write(archive, "[Content_Types].xml", $"""
                <?xml version="1.0"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
                  {worksheetType}
                </Types>
                """);
            Write(archive, "_rels/.rels", """
                <?xml version="1.0"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            string worksheet = includeWorksheet
                ? "<sheet name=\"Data\" sheetId=\"2\" r:id=\"rId2\"/>"
                : "";
            string names = includeWorksheet
                ? "<definedNames><definedName name=\"Local\" localSheetId=\"1\">Data!$A$1</definedName></definedNames>"
                : "";
            Write(archive, "xl/workbook.xml", $"""
                <?xml version="1.0"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Chart" sheetId="1" r:id="rId1"/>
                    {worksheet}
                  </sheets>
                  {names}
                </workbook>
                """);

            string worksheetRelationship = includeWorksheet
                ? "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                : "";
            Write(archive, "xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
                  {worksheetRelationship}
                </Relationships>
                """);
            Write(archive, "xl/chartsheets/sheet1.xml",
                """<?xml version="1.0"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""");
            if (includeWorksheet)
            {
                Write(archive, "xl/worksheets/sheet1.xml",
                    """<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");
            }
        }

        stream.Position = 0;
        return stream;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using Stream entry = archive.CreateEntry(name).Open();
        entry.Write(System.Text.Encoding.UTF8.GetBytes(content));
    }
}
