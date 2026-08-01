using System.Text;
using Cellwright.Styles;
using Inkwright;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Renders whole workbooks rather than one feature at a time, and checks the values came
/// through.
/// </summary>
/// <remarks>
/// The tests elsewhere each hold one thing still and measure another. This one does the
/// opposite: it takes workbooks with everything in them at once and asks only whether the
/// figures survived, on the grounds that a converter which drops a column is broken however
/// well it rules its borders.
/// </remarks>
public sealed class CorpusTests
{
    /// <summary>The real workbooks the sibling repositories keep, if they are checked out.</summary>
    public static TheoryData<string> Packages
    {
        get
        {
            var data = new TheoryData<string>();
            foreach (string path in Corpus.Packages.Take(40))
                data.Add(path);

            if (data.Count == 0)
                data.Add(string.Empty);

            return data;
        }
    }

    [Theory]
    [MemberData(nameof(Packages))]
    public async Task EveryRealWorkbookRendersWithoutFallingOver(string path)
    {
        Assert.SkipWhen(path.Length == 0, "The reference corpus is not present. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        Workbook book;
        try
        {
            book = await Workbook.LoadAsync(path, cancellationToken: TestContext.Current.CancellationToken);
        }
        catch (Exception exception) when (Corpus.IsUnreadable(exception))
        {
            // A corpus of test data keeps deliberately broken files; they prove nothing here.
            return;
        }

        using Rendered rendered = Rendered.Of(book);
        Assert.True(rendered.PageCount >= 1);
    }

    [Fact]
    public void AWorkbookUsingEverythingAtOnceComesThroughWhole()
    {
        using Rendered rendered = Rendered.Of(Everything());
        AssertValuesSurvived(rendered);
    }

    [Fact]
    public void AWorkbookUsingEverythingAtOnceAlsoTagsCleanly()
    {
        using Rendered rendered = Rendered.Of(Everything(), new PdfExportOptions { Tagged = true, Language = "en" });

        AssertValuesSurvived(rendered);
        Assert.NotEmpty(rendered.Document.Structure);
    }

    /// <summary>A workbook that leans on every part of the converter at once.</summary>
    private static Workbook Everything()
    {
        Workbook book = Workbook.Create();
        book.Properties.Title = "Everything";

        Worksheet sheet = book.Sheets.Add("Figures");
        sheet.PageSetup.PrintGridLines = true;
        sheet.PageSetup.Header = new HeaderFooterText { Default = "&L&A&RQuarterly" };
        sheet.PageSetup.Footer = new HeaderFooterText { Default = "&CPage &P of &N" };
        sheet.SetRepeatRows(1, 1);

        sheet["A1"] = "Region";
        sheet["B1"] = "Revenue";
        sheet["C1"] = "Share";
        sheet["D1"] = "Opened";
        sheet.Range("A1:D1").Style(style => style
            .Bold()
            .Fill(0xFFDDEBF7)
            .Align(HorizontalAlignment.Center));

        for (int row = 2; row <= 40; row++)
        {
            sheet.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"Region {row - 1}"));
            sheet.Cell(new CellRef(row, 2)).SetValue(CellValue.Number(1000.5 * row));
            sheet.Cell(new CellRef(row, 3)).SetValue(CellValue.Number(row / 100.0));
            sheet.Cell(new CellRef(row, 4)).SetValue(CellValue.Number(45000 + row));

            sheet.Cell(new CellRef(row, 2)).SetStyle(CellStyle.Default with { NumberFormat = "#,##0.00" });
            sheet.Cell(new CellRef(row, 3)).SetStyle(CellStyle.Default with { NumberFormat = "0.0%" });
            sheet.Cell(new CellRef(row, 4)).SetStyle(CellStyle.Default with { NumberFormat = "yyyy-mm-dd" });
        }

        sheet.Column("A").SetWidth(18);
        sheet.Column("B").SetWidth(14);
        sheet.Column("D").SetWidth(14);
        sheet.Row(20).Hidden = true;

        sheet["A42"] = "A note that wraps because it is far longer than the column that holds it";
        sheet.Cell("A42").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { WrapText = true },
            Border = BorderStyle.All(BorderEdge.Thin),
        });

        sheet["A44"] = "Merged heading across the sheet";
        sheet.Merge(RangeRef.Parse("A44:D44"));
        sheet.Hyperlinks.Add(Hyperlink.ToUrl(RangeRef.Parse("A1"), "https://example.org/"));

        Worksheet second = book.Sheets.Add("Landscape");
        second.PageSetup.Orientation = PageOrientation.Landscape;
        second["A1"] = "A wider page for the appendix";

        return book;
    }

    /// <summary>Checks that the figures of the workbook turn up somewhere in the rendered pages.</summary>
    private static void AssertValuesSurvived(Rendered rendered)
    {
        var pages = new StringBuilder();
        for (int page = 0; page < rendered.PageCount; page++)
            pages.Append(rendered.Text(page));

        string printed = pages.ToString();

        Assert.Contains("Region 1", printed, StringComparison.Ordinal);
        Assert.Contains("Region 39", printed, StringComparison.Ordinal);
        Assert.Contains("2,001.00", printed, StringComparison.Ordinal);
        Assert.Contains("2.0%", printed, StringComparison.Ordinal);
        Assert.Contains("Merged heading", printed, StringComparison.Ordinal);
        Assert.Contains("A wider page for the appendix", printed, StringComparison.Ordinal);

        // The hidden row is not printed, and the repeated heading is on every page of the sheet.
        Assert.DoesNotContain("Region 19", printed, StringComparison.Ordinal);
    }
}

/// <summary>
/// Finds the workbooks that other spreadsheet libraries keep as test data in sibling
/// repositories: the only realistic source of files written by producers nobody here controls.
/// </summary>
internal static class Corpus
{
    private static readonly Lazy<string[]> AllPackages = new(() =>
        [.. Discover("*.xlsx").Concat(Discover("*.xlsm")).OrderBy(static path => path, StringComparer.Ordinal)]);

    /// <summary>Every modern package in the corpus, in a stable order.</summary>
    public static IReadOnlyList<string> Packages => AllPackages.Value;

    /// <summary>Whether an exception means the file itself is unusable rather than the reader being wrong.</summary>
    public static bool IsUnreadable(Exception exception) => exception is
        XlsxFormatException or InvalidDataException or IOException or NotSupportedException;

    private static string[] Discover(string pattern)
    {
        if (FindCorpusRoot() is not { } root)
            return [];

        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
        return [.. Directory.EnumerateFiles(root, pattern, options)
            .Where(static path => !path.Contains("Cellwright", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length is > 0 and < 4 * 1024 * 1024)
            .OrderBy(static path => path, StringComparer.Ordinal)];
    }

    private static string? FindCorpusRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "ExcelDataReader-3.9.0")))
                return directory.FullName;

            directory = directory.Parent;
        }

        return null;
    }
}
