using System.IO.Compression;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// A zip declares how far each entry expands and is under no obligation to be honest. A package
/// built for the purpose turns a few kilobytes on disk into as much memory as it likes, which
/// for anything accepting files from strangers is the whole attack.
/// </summary>
public sealed class PackageLimitTests
{
    /// <summary>Builds a package whose shared strings expand far past what they cost on disk.</summary>
    private static MemoryStream Bomb(int megabytes, bool largeSharedStrings = false)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml",
                """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""" +
                """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""" +
                """<Default Extension="xml" ContentType="application/xml"/>""" +
                """<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>""" +
                """<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>""" +
                "</Types>");
            Write(archive, "_rels/.rels",
                """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
                """<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>""" +
                "</Relationships>");
            Write(archive, "xl/workbook.xml",
                """<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """ +
                """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""" +
                """<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>""");
            Write(archive, "xl/_rels/workbook.xml.rels",
                """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
                """<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>""" +
                """<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>""" +
                "</Relationships>");
            Write(archive, "xl/worksheets/sheet1.xml",
                """<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");

            if (largeSharedStrings)
            {
                ZipArchiveEntry strings = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.SmallestSize);
                using Stream stringContent = strings.Open();
                stringContent.Write(System.Text.Encoding.UTF8.GetBytes(
                    """<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>"""));
                byte[] text = new byte[1024 * 1024];
                Array.Fill(text, (byte)'x');
                for (int i = 0; i < megabytes; i++)
                    stringContent.Write(text);
                stringContent.Write("</t></si></sst>"u8);
            }

            // An unmodelled part, of the kind a real workbook fills with images: it is held in
            // memory whole so that saving can put it back. Highly compressible padding is what
            // keeps such a file small on disk while asking for as much memory as it likes.
            ZipArchiveEntry entry = archive.CreateEntry("xl/media/image1.bin", CompressionLevel.SmallestSize);
            using Stream content = entry.Open();
            byte[] chunk = new byte[1024 * 1024];
            for (int i = 0; i < megabytes; i++)
                content.Write(chunk);
        }

        stream.Position = 0;
        return stream;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using Stream entry = archive.CreateEntry(name).Open();
        entry.Write(System.Text.Encoding.UTF8.GetBytes(content));
    }

    [Fact]
    public async Task APartThatExpandsPastTheLimitIsRefused()
    {
        using MemoryStream package = Bomb(megabytes: 8);
        var options = new LoadOptions
        {
            LeaveOpen = true,
            Limits = new PackageLimits { MaxPartSize = 1024 * 1024 },
        };

        DocumentLoadLimitException failure = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await Workbook.LoadAsync(package, options, TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxPartBytes), failure.LimitName);
        Assert.Equal(1024 * 1024, failure.Limit);
    }

    [Fact]
    public async Task ASharedStringTableCannotBypassThePartLimitThroughItsStreamingReader()
    {
        using MemoryStream package = Bomb(megabytes: 8, largeSharedStrings: true);
        var options = new LoadOptions
        {
            LeaveOpen = true,
            Limits = new PackageLimits { MaxPartSize = 1024 * 1024 },
        };

        DocumentLoadLimitException failure = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await Workbook.LoadAsync(package, options, TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxPartBytes), failure.LimitName);
        Assert.Equal(1024 * 1024, failure.Limit);
    }

    [Fact]
    public async Task TheSameFileLoadsWhenTheLimitAllowsIt()
    {
        using MemoryStream package = Bomb(megabytes: 8);
        var options = new LoadOptions { LeaveOpen = true, Limits = PackageLimits.None };

        Workbook book = await Workbook.LoadAsync(package, options, TestContext.Current.CancellationToken);
        Assert.Single(book.Sheets);
    }

    /// <summary>
    /// The defaults have to be generous enough that no real workbook meets them; the corpus
    /// round-trip proves that at large, and this pins the numbers themselves.
    /// </summary>
    [Fact]
    public void TheDefaultsSitFarAboveAnyRealWorkbook()
    {
        Assert.Equal(512L * 1024 * 1024, PackageLimits.Default.MaxPartSize);
        Assert.Equal(2L * 1024 * 1024 * 1024, PackageLimits.Default.MaxTotalSize);
        Assert.Equal(long.MaxValue, PackageLimits.None.MaxPartSize);
    }
}
