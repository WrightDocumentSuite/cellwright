using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Where a note's box ends up. A note is two things in a package — an entry in the comments
/// part and a shape in a legacy drawing — and it is the shape that decides what the reader
/// sees, so the shape has to say which cells the box spans.
/// </summary>
public sealed class NoteGeometryTests
{
    [Fact]
    public async Task EveryNoteGetsItsOwnBoxNextToItsCell()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Notes");
        sheet.Comments.Add("B2", "first");
        sheet.Comments.Add("E10", "second");

        string vml = await ReadPartAsync(book, ".vml");

        // The anchor spans the two columns right of the cell and four rows from one above it.
        Assert.Contains("<x:Anchor>2, 15, 0, 2, 4, 15, 4, 16</x:Anchor>", vml, StringComparison.Ordinal);
        Assert.Contains("<x:Anchor>5, 15, 8, 2, 7, 15, 12, 16</x:Anchor>", vml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ANoteInTheFirstRowKeepsItsBoxOnTheSheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Notes");
        sheet.Comments.Add("A1", "top left");

        string vml = await ReadPartAsync(book, ".vml");

        Assert.Contains("<x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor>", vml, StringComparison.Ordinal);
    }

    /// <summary>
    /// Shape identifiers come out of a block per drawing, and the block a drawing owns is what
    /// its <c>idmap</c> declares. Two sheets claiming the same block is a file Excel repairs.
    /// </summary>
    [Fact]
    public async Task TwoSheetsWithNotesClaimDifferentShapeBlocks()
    {
        var book = Workbook.Create();
        book.Sheets.Add("First").Comments.Add("A1", "one");
        book.Sheets.Add("Second").Comments.Add("A1", "two");

        Dictionary<string, string> parts = await ReadAllPartsAsync(book);
        List<string> drawings = [.. parts
            .Where(static part => part.Key.EndsWith(".vml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static part => part.Key, StringComparer.Ordinal)
            .Select(static part => part.Value)];

        Assert.Equal(2, drawings.Count);
        Assert.Contains("data=\"1\"", drawings[0], StringComparison.Ordinal);
        Assert.Contains("_x0000_s1025", drawings[0], StringComparison.Ordinal);
        Assert.Contains("data=\"2\"", drawings[1], StringComparison.Ordinal);
        Assert.Contains("_x0000_s2049", drawings[1], StringComparison.Ordinal);
    }

    /// <summary>Serializes the workbook, validates it, and returns the part named this way.</summary>
    private static async Task<string> ReadPartAsync(Workbook book, string suffix)
    {
        Dictionary<string, string> parts = await ReadAllPartsAsync(book);
        return parts.First(part => part.Key.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)).Value;
    }

    private static async Task<Dictionary<string, string>> ReadAllPartsAsync(Workbook book)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        var parts = new Dictionary<string, string>(StringComparer.Ordinal);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true, Encoding.UTF8);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            parts[entry.FullName] = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
        }

        return parts;
    }
}
