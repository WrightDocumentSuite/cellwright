using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// Charts are the part of a real workbook this library models least and can least afford to
/// lose. They travel through byte for byte, which only holds while nothing disturbs the
/// relationships that reach them — and editing a sheet's notes rewrites relationships.
/// </summary>
public sealed class ChartPreservationTests
{
    /// <summary>A workbook that has both a chart and notes, which is what makes the case.</summary>
    private static string? FindChartedWorkbookWithNotes() => SampleCorpus.Files(int.MaxValue)
        .FirstOrDefault(static file => HasBoth(file));

    private static bool HasBoth(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);
            bool chart = false;
            bool comments = false;
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                chart |= entry.FullName.Contains("/charts/chart", StringComparison.OrdinalIgnoreCase);
                comments |= entry.FullName.Contains("comments", StringComparison.OrdinalIgnoreCase);
            }

            return chart && comments;
        }
        catch (Exception exception) when (SampleCorpus.IsUnreadable(exception))
        {
            return false;
        }
    }

    [Fact]
    public async Task EditingTheNotesOfAChartedSheetKeepsTheChart()
    {
        string? path = FindChartedWorkbookWithNotes();
        Assert.SkipWhen(path is null, "No sample workbook has both a chart and notes. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        HashSet<string> before = ChartParts(File.ReadAllBytes(path!));
        Assert.NotEmpty(before);

        Workbook book = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        Worksheet sheet = book.Sheets.First(static s => s.Comments.Count > 0);
        sheet.Comments.Add(sheet.Comments.First().Cell, "rewritten", "Tester");

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        HashSet<string> after = ChartParts(saved.ToArray());
        Assert.Equal(before, after);

        // A chart nothing points at is a chart Excel drops on the next save, so the drawing
        // that reaches it has to survive the rewritten relationships too.
        Assert.Contains(Relationships(saved.ToArray()), static target =>
            target.Contains("drawing", StringComparison.OrdinalIgnoreCase));

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("rewritten", reloaded.Sheets.First(static s => s.Comments.Count > 0).Comments.First().Text);
    }

    private static HashSet<string> ChartParts(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return
        [
            .. archive.Entries
                .Where(static entry => entry.FullName.Contains("/charts/", StringComparison.OrdinalIgnoreCase))
                .Select(static entry => entry.FullName),
        ];
    }

    /// <summary>Every relationship target in the package, so that a dangling chart shows up.</summary>
    private static List<string> Relationships(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var targets = new List<string>();
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                continue;

            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            targets.Add(reader.ReadToEnd());
        }

        return targets;
    }
}
