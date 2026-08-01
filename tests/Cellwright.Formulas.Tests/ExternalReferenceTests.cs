using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// References into another workbook (§18.14): the engine answers them from the values the
/// external link part cached, the way Excel answers them while the source file is closed.
/// </summary>
public sealed class ExternalReferenceTests
{
    [Fact]
    public async Task CachedValuesAnswerExternalReferences()
    {
        using MemoryStream package = BuildPackageWithExternalLink();
        Workbook book = await Workbook.LoadAsync(package, cancellationToken: TestContext.Current.CancellationToken);

        ExternalBook external = Assert.Single(book.ExternalBooks);
        Assert.Equal("prices.xlsx", external.TargetPath);
        Assert.Equal(["Data", "Rates Sheet"], external.SheetNames);
        Assert.Equal(7.5, external.GetCachedValue("Data", 2, 2).GetNumber());

        Worksheet sheet = book.Sheets[0];
        sheet.Cell("C1").SetFormula("[1]Data!B2");
        sheet.Cell("C2").SetFormula("SUM([1]Data!B2:B4)");
        sheet.Cell("C3").SetFormula("'[1]Rates Sheet'!A1");
        sheet.Cell("C4").SetFormula("[1]Nowhere!A1");
        sheet.Cell("C5").SetFormula("[1]Data!B2*10");
        book.Calculate();

        Assert.Equal(7.5, sheet["C1"].GetNumber());
        Assert.Equal(30.0, sheet["C2"].GetNumber());
        Assert.Equal("EUR", sheet["C3"].GetText());
        Assert.Equal(ExcelError.Reference, sheet["C4"].GetError());
        Assert.Equal(75.0, sheet["C5"].GetNumber());
    }

    [Fact]
    public void ABookWithoutLinksAnswersReferenceError()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("[1]Data!B2");
        book.Calculate();

        Assert.Equal(ExcelError.Reference, sheet["A1"].GetError());
    }

    /// <summary>
    /// A minimal package with one sheet and one external link whose cache holds a few values.
    /// </summary>
    private static MemoryStream BuildPackageWithExternalLink()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
                </Types>
                """);
            Write(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);
            Write(archive, "xl/workbook.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets>
                <externalReferences><externalReference r:id="rId2"/></externalReferences>
                </workbook>
                """);
            Write(archive, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
                </Relationships>
                """);
            Write(archive, "xl/worksheets/sheet1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
                </worksheet>
                """);
            Write(archive, "xl/externalLinks/externalLink1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <externalBook r:id="rId1">
                <sheetNames><sheetName val="Data"/><sheetName val="Rates Sheet"/></sheetNames>
                <sheetDataSet>
                <sheetData sheetId="0">
                <row r="2"><cell r="B2"><v>7.5</v></cell></row>
                <row r="3"><cell r="B3"><v>10</v></cell></row>
                <row r="4"><cell r="B4"><v>12.5</v></cell></row>
                </sheetData>
                <sheetData sheetId="1">
                <row r="1"><cell r="A1" t="str"><v>EUR</v></cell></row>
                </sheetData>
                </sheetDataSet>
                </externalBook>
                </externalLink>
                """);
            Write(archive, "xl/externalLinks/_rels/externalLink1.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="prices.xlsx" TargetMode="External"/>
                </Relationships>
                """);
        }

        stream.Position = 0;
        return stream;

        static void Write(ZipArchive archive, string path, string content)
        {
            using Stream entry = archive.CreateEntry(path).Open();
            entry.Write(Encoding.UTF8.GetBytes(content));
        }
    }
}
