using System.Text;
using Cellwright.IO;
using Cellwright.Xml;
using Xunit;

namespace Cellwright.Tests.IO;

public sealed class OpcPackageTests
{
    [Fact]
    public async Task WriteAndReadBack_PartsRelationshipsAndContentTypes()
    {
        using var stream = new MemoryStream();

        await using (OpcPackage package = await OpcPackage.CreateAsync(stream, leaveOpen: true, TestContext.Current.CancellationToken))
        {
            Stream part = await package.CreatePartAsync("/xl/workbook.xml", XlsxSchema.ContentTypeWorkbook, TestContext.Current.CancellationToken);
            await using (var writer = new Utf8XmlWriter(part))
            {
                writer.WriteDeclaration();
                writer.WriteRaw("<workbook/>"u8);
            }

            package.AddRelationship("/", new OpcRelationship("rId1", XlsxSchema.RelOfficeDocument, "xl/workbook.xml"));
            package.AddRelationship("/xl/workbook.xml", new OpcRelationship("rId1", XlsxSchema.RelWorksheet, "worksheets/sheet1.xml"));
        }

        stream.Position = 0;
        await using OpcPackage read = await OpcPackage.OpenReadAsync(stream, leaveOpen: true, TestContext.Current.CancellationToken);

        Assert.True(read.PartExists("/xl/workbook.xml"));
        Assert.True(read.PartExists("/[Content_Types].xml"));

        IReadOnlyList<OpcRelationship> rootRels = await read.GetRelationshipsAsync("/", TestContext.Current.CancellationToken);
        OpcRelationship officeDocument = Assert.Single(rootRels);
        Assert.Equal(XlsxSchema.RelOfficeDocument, officeDocument.Type);
        Assert.Equal("/xl/workbook.xml", OpcPath.Resolve("/", officeDocument.Target));

        IReadOnlyList<OpcRelationship> workbookRels = await read.GetRelationshipsAsync("/xl/workbook.xml", TestContext.Current.CancellationToken);
        Assert.Equal("/xl/worksheets/sheet1.xml", OpcPath.Resolve("/xl/workbook.xml", Assert.Single(workbookRels).Target));

        ContentTypeMap contentTypes = await read.GetContentTypesAsync(TestContext.Current.CancellationToken);
        Assert.Equal(XlsxSchema.ContentTypeWorkbook, contentTypes.GetContentType("/xl/workbook.xml"));
        Assert.Equal(XlsxSchema.ContentTypeRelationships, contentTypes.GetContentType("/_rels/.rels"));

        Stream reopened = await read.OpenPartAsync("/xl/workbook.xml", TestContext.Current.CancellationToken);
        await using (reopened)
        {
            using var readerText = new StreamReader(reopened, Encoding.UTF8);
            string xml = await readerText.ReadToEndAsync(TestContext.Current.CancellationToken);
            Assert.Contains("<workbook/>", xml, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task OpenPart_Missing_ThrowsXlsxFormatException()
    {
        using var stream = new MemoryStream();
        await using (OpcPackage package = await OpcPackage.CreateAsync(stream, leaveOpen: true, TestContext.Current.CancellationToken))
        {
            Stream part = await package.CreatePartAsync("/a.xml", "application/xml", TestContext.Current.CancellationToken);
            await part.DisposeAsync();
        }

        stream.Position = 0;
        await using OpcPackage read = await OpcPackage.OpenReadAsync(stream, leaveOpen: true, TestContext.Current.CancellationToken);
        await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await read.OpenPartAsync("/missing.xml", TestContext.Current.CancellationToken));
    }

    [Theory]
    [InlineData("/", "/_rels/.rels")]
    [InlineData("/xl/workbook.xml", "/xl/_rels/workbook.xml.rels")]
    [InlineData("/xl/worksheets/sheet1.xml", "/xl/worksheets/_rels/sheet1.xml.rels")]
    public void GetRelsPath_Works(string part, string expected)
    {
        Assert.Equal(expected, OpcPath.GetRelsPath(part));
    }

    [Theory]
    [InlineData("/xl/workbook.xml", "worksheets/sheet1.xml", "/xl/worksheets/sheet1.xml")]
    [InlineData("/xl/workbook.xml", "/xl/styles.xml", "/xl/styles.xml")]
    [InlineData("/xl/worksheets/sheet1.xml", "../styles.xml", "/xl/styles.xml")]
    [InlineData("/", "xl/workbook.xml", "/xl/workbook.xml")]
    [InlineData("/xl/workbook.xml", "./sharedStrings.xml", "/xl/sharedStrings.xml")]
    public void Resolve_Works(string source, string target, string expected)
    {
        Assert.Equal(expected, OpcPath.Resolve(source, target));
    }

    /// <summary>
    /// Part names differing only in the case of an ASCII letter are the same part; anything
    /// else is not (ECMA-376-2 6.2.2.3). The Kelvin sign is the trap: the general-purpose
    /// case-insensitive comparison in the runtime folds it onto the letter K, which would make
    /// two distinct parts collide and hand back the contents of whichever was stored second.
    /// </summary>
    [Theory]
    [InlineData("/xl/Workbook.xml", "/XL/workbook.XML", true)]
    [InlineData("/xl/workbook.xml", "/xl/workbook.xml", true)]
    [InlineData("/xl/K.xml", "/xl/k.xml", true)]
    [InlineData("/xl/\u212A.xml", "/xl/k.xml", false)]
    [InlineData("/xl/\u0130.xml", "/xl/i.xml", false)]
    [InlineData("/xl/\u00C9.xml", "/xl/\u00E9.xml", false)]
    public void PartNamesMatchOverAsciiOnly(string left, string right, bool same)
    {
        Assert.Equal(same, OpcNames.Same(left, right));
    }

    /// <summary>
    /// The extension is what follows the last dot of the last segment. A dot in a folder name
    /// belongs to the folder, and a part with no extension of its own has none.
    /// </summary>
    [Theory]
    [InlineData("/xl/workbook.xml", "xml")]
    [InlineData("/xl.old/workbook.xml", "xml")]
    [InlineData("/xl.old/workbook", null)]
    [InlineData("/workbook", null)]
    [InlineData("/xl/media/image1.PNG", "PNG")]
    public void TheExtensionComesFromTheLastSegment(string partPath, string? expected)
    {
        Assert.Equal(expected, ContentTypeMap.Extension(partPath));
    }

    /// <summary>
    /// A part covered by both a default and an override takes the override, and an extension
    /// matches whatever the file spelled it as.
    /// </summary>
    [Fact]
    public void AnOverrideBeatsTheExtensionDefault()
    {
        var map = new ContentTypeMap();
        map.AddDefault("xml", "application/xml");
        map.AddOverride("/xl/workbook.xml", XlsxSchema.ContentTypeWorkbook);

        Assert.Equal(XlsxSchema.ContentTypeWorkbook, map.GetContentType("/xl/workbook.xml"));
        Assert.Equal("application/xml", map.GetContentType("/xl/styles.XML"));
        Assert.Null(map.GetContentType("/xl.old/nameless"));
    }
}
