using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// Corners of ISO/IEC 29500 that a round trip used to quietly flatten: the third sheet state,
/// cell protection, the attributes of a formula element and the metadata indexes that tie a
/// cell to the part behind rich values.
/// </summary>
public sealed class SpecFidelityTests
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

    [Fact]
    public async Task VeryHiddenStaysVeryHidden()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Visible")["A1"] = 1;
        book.Sheets.Add("Hidden").Visibility = SheetVisibility.Hidden;
        book.Sheets.Add("Internal").Visibility = SheetVisibility.VeryHidden;

        Workbook loaded = await SaveAndReloadAsync(book);

        Assert.Equal(SheetVisibility.Visible, loaded.Sheets["Visible"].Visibility);
        Assert.Equal(SheetVisibility.Hidden, loaded.Sheets["Hidden"].Visibility);
        Assert.Equal(SheetVisibility.VeryHidden, loaded.Sheets["Internal"].Visibility);
        Assert.True(loaded.Sheets["Internal"].IsHidden);
    }

    [Fact]
    public async Task CellProtectionSurvivesAndClosedXmlAgrees()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Form");
        sheet["A1"] = "Label";
        sheet.Cell("B1").SetValue("input").SetStyle(CellStyle.Default.Unlocked());
        sheet.Cell("C1").SetValue(1).SetStyle(CellStyle.Default.HideFormula());
        sheet.Protection.Protect();

        Workbook reloaded = await SaveAndReloadAsync(book);
        Worksheet loaded = reloaded.Sheets["Form"];

        Assert.True(loaded.Cell("A1").StyleValue.Protection.Locked);
        Assert.False(loaded.Cell("B1").StyleValue.Protection.Locked);
        Assert.True(loaded.Cell("C1").StyleValue.Protection.Hidden);

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        using var interop = new XLWorkbook(stream);
        Assert.False(interop.Worksheet("Form").Cell("B1").Style.Protection.Locked);
    }

    [Fact]
    public async Task AlwaysRecalculateFlagOnAFormulaSurvives()
    {
        using var source = BuildWorkbookWithRawSheet("""
            <sheetData>
              <row r="1"><c r="A1"><f ca="1">RAND()</f><v>0.5</v></c></row>
            </sheetData>
            """);

        Workbook book = await Workbook.LoadAsync(source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        string saved = await SaveToTextAsync(book);

        Assert.Contains("ca=\"1\"", saved, StringComparison.Ordinal);
        Assert.Contains("RAND()", saved, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DataTableFormulaIsCarriedThroughAndReported()
    {
        using var source = BuildWorkbookWithRawSheet("""
            <sheetData>
              <row r="2">
                <c r="B2"><v>10</v></c>
                <c r="C2"><f t="dataTable" ref="C2:C3" dt2D="0" dtr="0" r1="A1"/><v>20</v></c>
              </row>
            </sheetData>
            """);

        var warnings = new List<WorkbookWarning>();
        var options = new LoadOptions { LeaveOpen = true, OnWarning = warnings.Add };
        Workbook book = await Workbook.LoadAsync(source, options, TestContext.Current.CancellationToken);

        Assert.Contains(warnings, w => w.Code == "DataTableFormula");

        string saved = await SaveToTextAsync(book);
        Assert.Contains("t=\"dataTable\"", saved, StringComparison.Ordinal);
        Assert.Contains("r1=\"A1\"", saved, StringComparison.Ordinal);
        Assert.Equal(20.0, book.Sheets[0]["C2"].GetNumber());
    }

    [Fact]
    public async Task CellMetadataIndexesAreCarriedThrough()
    {
        using var source = BuildWorkbookWithRawSheet("""
            <sheetData>
              <row r="1"><c r="A1" vm="3" cm="2" t="e"><v>#VALUE!</v></c></row>
            </sheetData>
            """);

        Workbook book = await Workbook.LoadAsync(source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        string saved = await SaveToTextAsync(book);

        Assert.Contains("vm=\"3\"", saved, StringComparison.Ordinal);
        Assert.Contains("cm=\"2\"", saved, StringComparison.Ordinal);
    }

    [Fact]
    public async Task InsertingRowsMovesTheMetadataWithTheCell()
    {
        using var source = BuildWorkbookWithRawSheet("""
            <sheetData>
              <row r="3"><c r="A3" vm="3"><v>1</v></c></row>
            </sheetData>
            """);

        Workbook book = await Workbook.LoadAsync(source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        book.Sheets[0].InsertRows(1, 2);

        string saved = await SaveToTextAsync(book);
        Assert.Contains("<c r=\"A5\" vm=\"3\"", saved, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RichTextKeepsItsRunsUntilTheCellIsRewritten()
    {
        using var source = BuildWorkbookWithRawSheet(
            """
            <sheetData>
              <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row>
            </sheetData>
            """,
            sharedStrings: """
                <si><r><t>plain </t></r><r><rPr><b/></rPr><t>bold</t></r></si>
                """);

        Workbook book = await Workbook.LoadAsync(source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("plain bold", book.Sheets[0]["A1"].GetText());

        string untouched = await ReadPartAsync(book, "sharedStrings.xml");
        Assert.Contains("<rPr>", untouched, StringComparison.Ordinal);
        Assert.Contains("<t>bold</t>", untouched, StringComparison.Ordinal);

        // Rewriting the text is a new string, and a new string has no formatting to keep.
        book.Sheets[0]["A1"] = "plain text";
        string edited = await ReadPartAsync(book, "sharedStrings.xml");
        Assert.Contains("<t>plain text</t>", edited, StringComparison.Ordinal);
    }

    [Fact]
    public async Task EqualRichTextWithDifferentRunsRemainsDistinct()
    {
        using var source = BuildWorkbookWithRawSheet(
            """
            <sheetData>
              <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
            </sheetData>
            """,
            sharedStrings: """
                <si><r><rPr><b/></rPr><t>same</t></r></si>
                <si><r><rPr><i/></rPr><t>same</t></r></si>
                """);

        Workbook book = await Workbook.LoadAsync(source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("same", book.Sheets[0]["A1"].GetText());
        Assert.Equal("same", book.Sheets[0]["B1"].GetText());

        string saved = await ReadPartAsync(book, "sharedStrings.xml");
        Assert.Contains("uniqueCount=\"2\"", saved, StringComparison.Ordinal);
        Assert.Contains("<b", saved, StringComparison.Ordinal);
        Assert.Contains("<i", saved, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RunsWrittenTheWayExcelWritesThemAreUnderstood()
    {
        // A run names its family rFont, where a style's font names it name. Everything else is
        // spelled the same, and a file from Excel carries the lot — so this is the case that
        // says the two spellings really do meet in one reader.
        using var source = BuildWorkbookWithRawSheet(
            """
            <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
            """,
            sharedStrings: """
                <si>
                  <r><t xml:space="preserve">plain </t></r>
                  <r>
                    <rPr><b/><i/><sz val="13"/><color rgb="FF9C0006"/><rFont val="Georgia"/><family val="1"/></rPr>
                    <t>dressed</t>
                  </r>
                </si>
                """);

        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        RichText? runs = book.Sheets[0].Cell("A1").RichText;

        Assert.NotNull(runs);
        Assert.Equal("plain dressed", runs.Text);
        Assert.Equal(2, runs.Count);

        Assert.Equal("plain ", runs[0].Text);
        Assert.Null(runs[0].Font);

        FontStyle dressed = runs[1].Font!;
        Assert.Equal("Georgia", dressed.Name);
        Assert.Equal(13, dressed.Size);
        Assert.True(dressed.Bold);
        Assert.True(dressed.Italic);
        Assert.Equal(1, dressed.Family);
        Assert.Equal(ExcelColor.FromRgb(0x9C, 0x00, 0x06), dressed.Color);
    }

    [Fact]
    public async Task PhoneticHintsDoNotBecomeRuns()
    {
        // A phonetic hint sits among the runs but annotates the text rather than being it.
        using var source = BuildWorkbookWithRawSheet(
            """
            <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
            """,
            sharedStrings: """
                <si>
                  <r><rPr><b/></rPr><t>漢字</t></r>
                  <rPh sb="0" eb="2"><t>カンジ</t></rPh>
                  <phoneticPr fontId="1"/>
                </si>
                """);

        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        RichText? runs = book.Sheets[0].Cell("A1").RichText;

        Assert.NotNull(runs);
        Assert.Single(runs);
        Assert.Equal("漢字", runs[0].Text);
        Assert.True(runs[0].Font!.Bold);
    }

    [Fact]
    public async Task StrictFilesOpenAndAnnounceThatTheyDoNotRoundTrip()
    {
        string? path = SampleCorpus.Files(int.MaxValue)
            .FirstOrDefault(static file => file.EndsWith("-strict.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "No strict-conformance sample workbook was found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        var warnings = new List<WorkbookWarning>();
        var options = new LoadOptions { OnWarning = warnings.Add };
        Workbook book = await Workbook.LoadAsync(path!, options, TestContext.Current.CancellationToken);

        Assert.NotEmpty(book.Sheets);
        Assert.Contains(warnings, w => w.Code == "StrictConformance");
    }

    [Fact]
    public async Task ARequiredExtensionIsReported()
    {
        using var source = BuildWorkbookWithRawSheet("<sheetData/>", mustUnderstand: "x15");

        var warnings = new List<WorkbookWarning>();
        var options = new LoadOptions { LeaveOpen = true, OnWarning = warnings.Add };
        await Workbook.LoadAsync(source, options, TestContext.Current.CancellationToken);

        Assert.Contains(warnings, w => w.Code == "MustUnderstand" && w.Message.Contains("x15", StringComparison.Ordinal));
    }

    /// <summary>Serializes the workbook and returns the text of the part whose name ends this way.</summary>
    private static async Task<string> ReadPartAsync(Workbook book, string suffix)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.Entries.First(e => e.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
        using var reader = new StreamReader(entry.Open());
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }

    /// <summary>Serializes the workbook and returns the text of its first worksheet part.</summary>
    private static async Task<string> SaveToTextAsync(Workbook book)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.Entries.First(static e =>
            e.FullName.Contains("worksheets/", StringComparison.OrdinalIgnoreCase));
        using var reader = new StreamReader(entry.Open());
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }

    /// <summary>
    /// Builds the smallest package that holds a hand-written worksheet body, which is how these
    /// tests get at markup no producer of ours would ever generate.
    /// </summary>
    private static MemoryStream BuildWorkbookWithRawSheet(
        string sheetData, string? sharedStrings = null, string? mustUnderstand = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                {(sharedStrings is null ? "" : "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>")}
                </Types>
                """);

            Write(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            Write(archive, "xl/workbook.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"{(mustUnderstand is null ? "" : $" mc:MustUnderstand=\"{mustUnderstand}\"")}>
                <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);

            Write(archive, "xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                {(sharedStrings is null ? "" : "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>")}
                </Relationships>
                """);

            if (sharedStrings is not null)
            {
                Write(archive, "xl/sharedStrings.xml", $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
                    {sharedStrings}
                    </sst>
                    """);
            }

            Write(archive, "xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                {sheetData}
                </worksheet>
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
