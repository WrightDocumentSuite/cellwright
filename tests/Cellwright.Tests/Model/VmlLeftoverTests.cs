using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Form controls share the legacy VML part with the comment boxes. Editing a note regenerates
/// the part, and the shapes that are not notes have to come along — a checkbox is not a fair
/// price for fixing a typo in a comment.
/// </summary>
public sealed class VmlLeftoverTests
{
    private const string CheckboxVml =
        "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\">" +
        "<v:stroke joinstyle=\"miter\"/><v:path shadowok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/>" +
        "</v:shapetype>" +
        "<v:shape id=\"_x0000_s2052\" type=\"#_x0000_t201\" style=\"position:absolute;margin-left:100pt;margin-top:10pt;" +
        "width:60pt;height:15pt;z-index:2\" o:insetmode=\"auto\">" +
        "<x:ClientData ObjectType=\"Checkbox\"><x:Anchor>2, 0, 1, 0, 4, 0, 2, 0</x:Anchor><x:Checked>1</x:Checked></x:ClientData>" +
        "</v:shape>";

    [Fact]
    public async Task EditingComments_KeepsFormControlShapes()
    {
        using MemoryStream package = await BuildBookWithCommentAndCheckboxAsync();
        Workbook book = await Workbook.LoadAsync(package, cancellationToken: TestContext.Current.CancellationToken);

        Worksheet sheet = book.Sheets[0];
        sheet.Comments.Add("B2", "второе примечание", author: "Ревизор");

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string vml = ReadReferencedVml(saved);
        Assert.Contains("ObjectType=\"Checkbox\"", vml, StringComparison.Ordinal);
        Assert.Contains("_x0000_s2052", vml, StringComparison.Ordinal);
        Assert.Contains("_x0000_t201", vml, StringComparison.Ordinal);
        Assert.Contains("ObjectType=\"Note\"", vml, StringComparison.Ordinal);

        // The notes are renumbered past the kept shape, so everything sits in its block
        // (2052 / 1024 = 2) and the map declares it.
        Assert.Contains("data=\"2\"", vml, StringComparison.Ordinal);

        // The regenerated notes must not collide with the kept shape's id.
        Assert.DoesNotContain("id=\"_x0000_s2052\" type=\"#_x0000_t202\"", vml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RemovingEveryComment_StillKeepsFormControlShapes()
    {
        using MemoryStream package = await BuildBookWithCommentAndCheckboxAsync();
        Workbook book = await Workbook.LoadAsync(package, cancellationToken: TestContext.Current.CancellationToken);

        Worksheet sheet = book.Sheets[0];
        sheet.Comments.Remove(CellRef.Parse("A1"));
        Assert.Empty(sheet.Comments);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string vml = ReadReferencedVml(saved);
        Assert.Contains("ObjectType=\"Checkbox\"", vml, StringComparison.Ordinal);
        Assert.DoesNotContain("ObjectType=\"Note\"", vml, StringComparison.Ordinal);

        // With the notes gone the sheet stops referencing a comments part; the VML relationship
        // alone stays for the control.
        string rels = ReadEntry(saved, "xl/worksheets/_rels/sheet1.xml.rels");
        Assert.DoesNotContain("comments", rels, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("vmlDrawing", rels, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UntouchedComments_KeepTheOriginalVmlVerbatim()
    {
        using MemoryStream package = await BuildBookWithCommentAndCheckboxAsync();
        Workbook book = await Workbook.LoadAsync(package, cancellationToken: TestContext.Current.CancellationToken);
        book.Sheets[0]["C3"] = 42; // Any edit that does not touch the notes.

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string vml = ReadReferencedVml(saved);
        Assert.Contains("ObjectType=\"Checkbox\"", vml, StringComparison.Ordinal);
        Assert.Contains("ObjectType=\"Note\"", vml, StringComparison.Ordinal);
    }

    /// <summary>A package with one note on A1 whose VML also hosts a checkbox control.</summary>
    private static async Task<MemoryStream> BuildBookWithCommentAndCheckboxAsync()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet.Comments.Add("A1", "проверить", author: "Ревизор");

        var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        (string vmlEntry, string vml) = FindVmlEntry(stream);
        vml = vml.Replace("</xml>", CheckboxVml + "</xml>", StringComparison.Ordinal);
        return RewriteEntry(stream, vmlEntry, vml);
    }

    /// <summary>The VML part the first sheet actually points at — orphans do not count.</summary>
    private static string ReadReferencedVml(MemoryStream package)
    {
        string rels = ReadEntry(package, "xl/worksheets/_rels/sheet1.xml.rels");
        int target = rels.IndexOf("drawings/vmlDrawing", StringComparison.Ordinal);
        Assert.True(target > 0, "the sheet references no VML part");
        int end = rels.IndexOf('"', target);
        return ReadEntry(package, "xl/" + rels[target..end]);
    }

    private static string ReadEntry(MemoryStream package, string entryName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(archive.GetEntry(entryName)!.Open());
        return reader.ReadToEnd();
    }

    private static (string Name, string Content) FindVmlEntry(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.Entries.First(static e =>
            e.FullName.EndsWith(".vml", StringComparison.OrdinalIgnoreCase));
        using var reader = new StreamReader(entry.Open());
        return (entry.FullName, reader.ReadToEnd());
    }

    private static MemoryStream RewriteEntry(MemoryStream package, string entryName, string content)
    {
        var rewritten = new MemoryStream();
        package.Position = 0;
        using (var source = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                using Stream output = target.CreateEntry(entry.FullName).Open();
                if (entry.FullName == entryName)
                {
                    output.Write(Encoding.UTF8.GetBytes(content));
                }
                else
                {
                    using Stream input = entry.Open();
                    input.CopyTo(output);
                }
            }
        }

        rewritten.Position = 0;
        return rewritten;
    }
}
