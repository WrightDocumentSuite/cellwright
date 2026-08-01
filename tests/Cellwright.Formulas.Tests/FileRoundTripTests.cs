using System.IO.Compression;
using ClosedXML.Excel;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// What happens when formulas travel through a real file: shared formulas, array markers,
/// defined names and the recalculation flag Excel reads on open.
/// </summary>
public sealed class FileRoundTripTests
{
    [Fact]
    public async Task SharedFormulas_AreExpandedIntoRealFormulas()
    {
        using var stream = new MemoryStream();
        WriteSharedFormulaWorkbook(stream);
        stream.Position = 0;

        Workbook book = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Worksheet sheet = book.Sheets[0];

        // Before expansion only the master carries text; the engine restores the rest.
        Assert.Equal("A1*2", sheet.Cell("B1").Formula);
        Assert.Null(sheet.Cell("B2").Formula);

        book.Calculate();

        Assert.Equal("A2*2", sheet.Cell("B2").Formula);
        Assert.Equal("A3*2", sheet.Cell("B3").Formula);
        Assert.Equal(2.0, sheet["B1"].GetNumber());
        Assert.Equal(4.0, sheet["B2"].GetNumber());
        Assert.Equal(6.0, sheet["B3"].GetNumber());
    }

    [Fact]
    public async Task ArrayFormulaMarker_SurvivesLoadAndSave()
    {
        using var source = new MemoryStream();
        WriteArrayFormulaWorkbook(source);
        source.Position = 0;

        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        saved.Position = 0;

        string sheetXml = ReadEntry(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("t=\"array\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("ref=\"C1:C2\"", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DefinedNames_RoundTripThroughTheFile()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        Worksheet other = book.Sheets.Add("Other");
        sheet["A1"] = 5;
        sheet["A2"] = 7;
        book.Names.Add("Totals", "Data!$A$1:$A$2");
        book.Names.Add("Local", "Other!$B$1", other);
        sheet.Cell("B1").SetFormula("SUM(Totals)");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        Workbook reloaded = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Assert.Equal(2, reloaded.Names.Count);
        Assert.Equal("Data!$A$1:$A$2", reloaded.Names.Resolve("Totals", null)!.RefersTo);
        Assert.Equal(reloaded.Sheets["Other"], reloaded.Names.Find("Local", reloaded.Sheets["Other"])!.Scope);

        reloaded.Calculate();
        Assert.Equal(12.0, reloaded.Sheets["Data"]["B1"].GetNumber());
    }

    [Fact]
    public async Task CalculatedWorkbook_DropsTheFullRecalculationFlag()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 6;
        sheet.Cell("A2").SetFormula("A1*7");
        book.Calculate();

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        string workbookXml = ReadEntry(stream, "xl/workbook.xml");
        Assert.DoesNotContain("fullCalcOnLoad", workbookXml, StringComparison.Ordinal);

        // The cached value travels with the formula, so other readers see it without calculating.
        stream.Position = 0;
        using var cross = new XLWorkbook(stream);
        Assert.Equal(42.0, cross.Worksheet("S").Cell("A2").GetDouble());
    }

    [Fact]
    public async Task UncalculatedWorkbook_KeepsTheFullRecalculationFlag()
    {
        var book = Workbook.Create();
        book.Sheets.Add("S").Cell("A1").SetFormula("1+1");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        Assert.Contains("fullCalcOnLoad", ReadEntry(stream, "xl/workbook.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task FormulasWrittenByCellwright_AreReadBackByClosedXml()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 3;
        sheet["A2"] = 4;
        sheet.Cell("A3").SetFormula("SUM(A1:A2)");
        sheet.Cell("A4").SetFormula("IF(A3>5,\"big\",\"small\")");
        book.Calculate();

        Assert.Equal(7.0, sheet["A3"].GetNumber());
        Assert.Equal("big", sheet["A4"].GetText());

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var cross = new XLWorkbook(stream);
        IXLWorksheet crossSheet = cross.Worksheet("S");
        Assert.Equal("SUM(A1:A2)", crossSheet.Cell("A3").FormulaA1);
        Assert.Equal(7.0, crossSheet.Cell("A3").GetDouble());
        Assert.Equal("big", crossSheet.Cell("A4").GetString());
    }

    private static string ReadEntry(Stream package, string entryName)
    {
        package.Position = 0;
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(zip.GetEntry(entryName)!.Open());
        return reader.ReadToEnd();
    }

    private static void WriteSharedFormulaWorkbook(Stream stream) => WriteWorkbook(
        stream,
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><v>1</v></c>
              <c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>
            </row>
            <row r="2">
              <c r="A2"><v>2</v></c>
              <c r="B2"><f t="shared" si="0"/><v>4</v></c>
            </row>
            <row r="3">
              <c r="A3"><v>3</v></c>
              <c r="B3"><f t="shared" si="0"/><v>6</v></c>
            </row>
          </sheetData>
        </worksheet>
        """);

    private static void WriteArrayFormulaWorkbook(Stream stream) => WriteWorkbook(
        stream,
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><v>2</v></c>
              <c r="C1"><f t="array" ref="C1:C2">A1*2</f><v>4</v></c>
            </row>
            <row r="2"><c r="A2"><v>3</v></c></row>
          </sheetData>
        </worksheet>
        """);

    private static void WriteWorkbook(Stream stream, string sheetXml)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        Write(zip, "[Content_Types].xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);
        Write(zip, "_rels/.rels",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);
        Write(zip, "xl/_rels/workbook.xml.rels",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """);
        Write(zip, "xl/workbook.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
            </workbook>
            """);
        Write(zip, "xl/worksheets/sheet1.xml", sheetXml);
    }

    private static void Write(ZipArchive zip, string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(content);
    }
}
