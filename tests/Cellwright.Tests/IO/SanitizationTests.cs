using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Cellwright.IO;
using Cellwright.Sanitization;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// A sanitized copy must agree with itself, and a workbook has more ways to refer to a link than
/// a relationship: formulas, defined names and rules all spell it as an index in their text.
/// </summary>
public sealed class SanitizationTests
{
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Sheet = "xl/worksheets/sheet1.xml";
    private const string Book = "xl/workbook.xml";
    private const string BookRels = "xl/_rels/workbook.xml.rels";
    private const string SheetRels = "xl/worksheets/_rels/sheet1.xml.rels";

    private static async Task<byte[]> BasePackageAsync()
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = "Ordinary";
        using var buffer = new MemoryStream();
        await workbook.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        return buffer.ToArray();
    }

    private static string Read(ZipArchive archive, string name)
    {
        using var reader = new StreamReader(archive.GetEntry(name)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static void Put(ZipArchive archive, string name, string content)
    {
        archive.GetEntry(name)?.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static void AddRelationship(ZipArchive archive, string relsPath, string id, string type, string target, bool external = false)
    {
        string relationship = $"<Relationship Id=\"{id}\" Type=\"{type}\" Target=\"{target}\"{(external ? " TargetMode=\"External\"" : string.Empty)}/>";
        string current = archive.GetEntry(relsPath) is null
            ? $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{PackageRelationships}\"></Relationships>"
            : Read(archive, relsPath);
        Put(archive, relsPath, current.Replace("</Relationships>", relationship + "</Relationships>", StringComparison.Ordinal));
    }

    private static void AddContentType(ZipArchive archive, string declaration)
    {
        string types = Read(archive, "[Content_Types].xml");
        Put(archive, "[Content_Types].xml", types.Replace("</Types>", declaration + "</Types>", StringComparison.Ordinal));
    }

    private static void Replace(ZipArchive archive, string name, string from, string to) =>
        Put(archive, name, Read(archive, name).Replace(from, to, StringComparison.Ordinal));

    private static void DeclareRelationshipPrefix(ZipArchive archive, string name)
    {
        string text = Read(archive, name);
        if (text.Contains("xmlns:r=", StringComparison.Ordinal))
            return;

        int root = text.IndexOf(name == Book ? "<workbook" : "<worksheet", StringComparison.Ordinal);
        int end = text.IndexOf('>', root);
        Put(archive, name, text.Insert(end, $" xmlns:r=\"{Relationships}\""));
    }

    private static byte[] Edit(byte[] package, Action<ZipArchive> edit)
    {
        using var buffer = new MemoryStream();
        buffer.Write(package);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
            edit(archive);

        return buffer.ToArray();
    }

    private static Task<byte[]> WithExternalWorkbookAsync(string formula = "[1]Rates!A1*2") =>
        WithAsync(archive =>
        {
            DeclareRelationshipPrefix(archive, Book);
            Put(archive, "xl/externalLinks/externalLink1.xml", "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
            AddContentType(archive, "<Override PartName=\"/xl/externalLinks/externalLink1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml\"/>");
            AddRelationship(archive, "xl/externalLinks/_rels/externalLink1.xml.rels", "rId1", Relationships + "/externalLinkPath", "file:///C:/Secret/rates.xlsx", external: true);
            AddRelationship(archive, BookRels, "rIdExt", Relationships + "/externalLink", "externalLinks/externalLink1.xml");
            Replace(archive, Book, "</sheets>", "</sheets><externalReferences><externalReference r:id=\"rIdExt\"/></externalReferences><definedNames><definedName name=\"Linked\">[1]Rates!$A$1</definedName><definedName name=\"Local\">Data!$A$1</definedName></definedNames>");
            Replace(archive, Sheet, "</sheetData>", $"<row r=\"5\"><c r=\"B5\"><f>{formula}</f><v>42</v></c><c r=\"C5\"><f>Sales[2024]</f><v>7</v></c></row></sheetData>");
        });

    private static async Task<byte[]> WithAsync(Action<ZipArchive> edit) => Edit(await BasePackageAsync(), edit);

    private static (byte[] Package, SanitizeReport Report) Run(byte[] package, ContentCategories categories) =>
        new XlsxContentSanitizer(categories).Run(package);

    private static HashSet<string> Names(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return new HashSet<string>(archive.Entries.Select(static e => e.FullName), StringComparer.OrdinalIgnoreCase);
    }

    private static string Entry(byte[] package, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return Read(archive, name);
    }

    private static void AssertPackageAgrees(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        HashSet<string> names = new(archive.Entries.Select(static e => OpcPath.ToPartPath(e.FullName)), StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries.Where(static e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToArray())
        {
            string source = OpcPath.GetSourcePart(OpcPath.ToPartPath(entry.FullName));
            foreach (XElement relationship in XDocument.Parse(Read(archive, entry.FullName)).Root!.Elements())
            {
                if ((string?)relationship.Attribute("TargetMode") == "External")
                    continue;

                string target = OpcPath.Resolve(source, (string)relationship.Attribute("Target")!);
                Assert.True(names.Contains(target), $"{entry.FullName} points at {target}, which is not in the package.");
            }
        }

        foreach (XElement entry in XDocument.Parse(Read(archive, "[Content_Types].xml")).Root!.Elements().Where(static e => e.Name.LocalName == "Override"))
            Assert.True(names.Contains((string)entry.Attribute("PartName")!), $"An override names {entry.Attribute("PartName")}, which is not in the package.");
    }

    [Fact]
    public async Task ALinkedFormulaKeepsItsValueAndLosesTheLink()
    {
        (byte[] package, SanitizeReport report) = Run(await WithExternalWorkbookAsync(), ContentCategories.ExternalLinks);

        Assert.DoesNotContain("xl/externalLinks/externalLink1.xml", Names(package));
        string sheet = Entry(package, Sheet);
        Assert.DoesNotContain("[1]Rates", sheet, StringComparison.Ordinal);
        Assert.Contains("<v>42</v>", sheet, StringComparison.Ordinal);

        string book = Entry(package, Book);
        Assert.DoesNotContain("externalReference", book, StringComparison.Ordinal);
        Assert.DoesNotContain("Linked", book, StringComparison.Ordinal);
        Assert.Contains("Local", book, StringComparison.Ordinal);
        Assert.Contains(report.Removed, static item => item.Detail.Contains("cached values", StringComparison.Ordinal));
        AssertPackageAgrees(package);

        Workbook reloaded = await Workbook.LoadAsync(new MemoryStream(package), new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Null(reloaded.Sheets["Data"].Cell("B5").Formula);
        Assert.Equal(42.0, reloaded.Sheets["Data"]["B5"].GetNumber());
    }

    [Fact]
    public async Task AStructuredReferenceToANumberedColumnIsNotALink()
    {
        (byte[] package, _) = Run(await WithExternalWorkbookAsync(), ContentCategories.ExternalLinks);

        Assert.Contains("<f>Sales[2024]</f>", Entry(package, Sheet), StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnUnselectedLinkIsLeftAndReported()
    {
        (byte[] package, SanitizeReport report) = Run(await WithExternalWorkbookAsync(), ContentCategories.Macros);

        Assert.Contains("xl/externalLinks/externalLink1.xml", Names(package));
        Assert.Contains("[1]Rates", Entry(package, Sheet), StringComparison.Ordinal);
        Assert.Empty(report.Removed);
        Assert.Contains(report.Retained, static item => item.Category == ContentCategories.ExternalLinks);
    }

    [Fact]
    public async Task AHyperlinkIsRemovedAndItsEmptyContainerWithIt()
    {
        byte[] input = await WithAsync(static archive =>
        {
            DeclareRelationshipPrefix(archive, Sheet);
            AddRelationship(archive, SheetRels, "rIdLink", Relationships + "/hyperlink", "https://tracker.example/x", external: true);
            Replace(archive, Sheet, "</sheetData>", "</sheetData><hyperlinks><hyperlink ref=\"A1\" r:id=\"rIdLink\"/></hyperlinks>");
        });

        (byte[] package, _) = Run(input, ContentCategories.ExternalLinks);

        Assert.DoesNotContain("hyperlink", Entry(package, Sheet), StringComparison.Ordinal);
        Assert.Contains("Ordinary", Entry(package, "xl/sharedStrings.xml") + Entry(package, Sheet), StringComparison.Ordinal);
        AssertPackageAgrees(package);
    }

    [Fact]
    public async Task AnOleObjectGoesWithItsAlternateContentAndPreview()
    {
        byte[] input = await WithAsync(static archive =>
        {
            DeclareRelationshipPrefix(archive, Sheet);
            Put(archive, "xl/embeddings/oleObject1.bin", "payload");
            Put(archive, "xl/media/preview.png", "preview");
            AddContentType(archive, "<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.oleObject\"/>");
            AddContentType(archive, "<Default Extension=\"png\" ContentType=\"image/png\"/>");
            AddRelationship(archive, SheetRels, "rIdOle", Relationships + "/oleObject", "../embeddings/oleObject1.bin");
            AddRelationship(archive, SheetRels, "rIdPrev", Relationships + "/image", "../media/preview.png");
            Replace(
                archive,
                Sheet,
                "</sheetData>",
                "</sheetData><oleObjects><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                "<mc:Choice Requires=\"x14\"><oleObject progId=\"P\" shapeId=\"1\" r:id=\"rIdOle\"><objectPr r:id=\"rIdPrev\"/></oleObject></mc:Choice>" +
                "<mc:Fallback><oleObject progId=\"P\" shapeId=\"1\" r:id=\"rIdOle\"/></mc:Fallback></mc:AlternateContent></oleObjects>");
        });

        (byte[] package, SanitizeReport report) = Run(input, ContentCategories.EmbeddedObjects);

        Assert.DoesNotContain("oleObject", Entry(package, Sheet), StringComparison.Ordinal);
        Assert.DoesNotContain("AlternateContent", Entry(package, Sheet), StringComparison.Ordinal);
        Assert.DoesNotContain("xl/embeddings/oleObject1.bin", Names(package));
        Assert.DoesNotContain("xl/media/preview.png", Names(package));
        Assert.Contains(report.Removed, static item => item is { Kind: SanitizedItemKind.Part, Category: ContentCategories.EmbeddedObjects });
        AssertPackageAgrees(package);
    }

    [Fact]
    public async Task VbaGoesAndTheWorkbookStopsClaimingToHaveIt()
    {
        byte[] input = await WithAsync(static archive =>
        {
            Put(archive, "xl/vbaProject.bin", "vba");
            AddContentType(archive, "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
            AddRelationship(archive, BookRels, "rIdVba", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin");
            Replace(archive, "[Content_Types].xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
        });

        (byte[] package, _) = Run(input, ContentCategories.Macros);

        Assert.DoesNotContain("xl/vbaProject.bin", Names(package));
        string types = Entry(package, "[Content_Types].xml");
        Assert.DoesNotContain("macroEnabled", types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject", types, StringComparison.Ordinal);
        Assert.Contains("spreadsheetml.sheet.main+xml", types, StringComparison.Ordinal);
        AssertPackageAgrees(package);
    }

    [Fact]
    public async Task RemovingAMacroSheetShiftsTheIndexesThatFollowIt()
    {
        byte[] input = await WithAsync(static archive =>
        {
            DeclareRelationshipPrefix(archive, Book);
            Put(archive, "xl/macrosheets/sheet1.xml", "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/>");
            AddContentType(archive, "<Override PartName=\"/xl/macrosheets/sheet1.xml\" ContentType=\"application/vnd.ms-excel.macrosheet+xml\"/>");
            AddRelationship(archive, BookRels, "rIdMacro", "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "macrosheets/sheet1.xml");
            string book = Read(archive, Book);
            int first = book.IndexOf("<sheets>", StringComparison.Ordinal) + "<sheets>".Length;
            book = book.Insert(first, "<sheet name=\"Auto_Open\" sheetId=\"99\" r:id=\"rIdMacro\"/>");
            book = book.Replace("</sheets>", "</sheets><definedNames><definedName name=\"Inner\" localSheetId=\"0\">Auto_Open!$A$1</definedName><definedName name=\"Kept\" localSheetId=\"1\">Data!$A$1</definedName></definedNames>", StringComparison.Ordinal);
            Put(archive, Book, book);
        });

        (byte[] package, _) = Run(input, ContentCategories.Macros);

        string result = Entry(package, Book);
        Assert.DoesNotContain("Auto_Open", result, StringComparison.Ordinal);
        Assert.DoesNotContain("Inner", result, StringComparison.Ordinal);
        Assert.Contains("name=\"Kept\" localSheetId=\"0\"", result, StringComparison.Ordinal);
        Assert.DoesNotContain("xl/macrosheets/sheet1.xml", Names(package));
        AssertPackageAgrees(package);
    }

    [Fact]
    public async Task ACalculationChainDoesNotOutliveTheFormulasItLists()
    {
        byte[] input = Edit(await WithExternalWorkbookAsync(), static archive =>
        {
            Put(archive, "xl/calcChain.xml", "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"B5\" i=\"1\"/></calcChain>");
            AddContentType(archive, "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>");
            AddRelationship(archive, BookRels, "rIdCalc", Relationships + "/calcChain", "calcChain.xml");
        });

        (byte[] package, _) = Run(input, ContentCategories.ExternalLinks);

        Assert.DoesNotContain("xl/calcChain.xml", Names(package));
        AssertPackageAgrees(package);
    }

    [Fact]
    public async Task SavingFromAWorkbookLeavesTheWorkbookUnchanged()
    {
        byte[] input = await WithExternalWorkbookAsync();
        Workbook workbook = await Workbook.LoadAsync(new MemoryStream(input), new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using var output = new MemoryStream();
        SanitizeReport report = await workbook.SaveSanitizedAsync(output, ContentCategories.ExternalLinks, TestContext.Current.CancellationToken);

        Assert.DoesNotContain("xl/externalLinks/externalLink1.xml", Names(output.ToArray()));
        Assert.NotEmpty(report.Removed);
        Assert.NotNull(workbook.Sheets["Data"].Cell("B5").Formula);
    }

    [Fact]
    public async Task ACleanWorkbookComesThroughUnharmed()
    {
        (byte[] package, SanitizeReport report) = Run(await BasePackageAsync(), ContentCategories.All);

        Assert.Empty(report.Removed);
        Assert.Empty(report.Retained);
        AssertPackageAgrees(package);

        Workbook reloaded = await Workbook.LoadAsync(new MemoryStream(package), new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("Ordinary", reloaded.Sheets["Data"]["A1"].GetText());
    }
}
