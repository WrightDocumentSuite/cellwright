using System.IO.Compression;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>Part-name equivalence is ASCII-only throughout both reading and preservation.</summary>
public sealed class OpcPreservationNameTests
{
    [Fact]
    public async Task UnicodeCaseFoldingDoesNotMergeDistinctPreservedParts()
    {
        using MemoryStream source = Package();
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await book.SaveAsync(
            saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, StringComparer.Ordinal);
        Assert.Equal("latin K", Read(parts["xl/K.xml"]));
        Assert.Equal("kelvin sign", Read(parts["xl/\u212A.xml"]));
    }

    private static MemoryStream Package()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """
                <?xml version="1.0"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """);
            Write(archive, "_rels/.rels", """
                <?xml version="1.0"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);
            Write(archive, "xl/workbook.xml", """
                <?xml version="1.0"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);
            Write(archive, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """);
            Write(archive, "xl/worksheets/sheet1.xml",
                """<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");
            Write(archive, "xl/K.xml", "latin K");
            Write(archive, "xl/\u212A.xml", "kelvin sign");
        }

        stream.Position = 0;
        return stream;
    }

    private static string Read(ZipArchiveEntry entry)
    {
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using Stream entry = archive.CreateEntry(name).Open();
        entry.Write(System.Text.Encoding.UTF8.GetBytes(content));
    }
}
