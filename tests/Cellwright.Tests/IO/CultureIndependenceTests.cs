using System.Globalization;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// A spreadsheet is culture-neutral: every number in it is written with a full stop and read
/// back the same way, whoever opens it. That holds only while every parse and every format says
/// so, and one call that forgets writes a file that is wrong in Berlin and right in Boston.
/// The build runs the whole suite twice, the second time under a culture whose decimal
/// separator is a comma, so a slip has somewhere to show up.
/// </summary>
public sealed class CultureIndependenceTests
{
    /// <summary>
    /// Proves the second run is actually a second run. Without this the culture leg of the
    /// build could stop applying anything and still pass, which is worse than not having it.
    /// </summary>
    [Fact]
    public void TheRequestedCultureIsTheOneInEffect()
    {
        string? requested = Environment.GetEnvironmentVariable("CELLWRIGHT_TEST_CULTURE");
        Assert.SkipWhen(string.IsNullOrEmpty(requested), "The suite is running under the default culture.");

        Assert.Equal(requested, CultureInfo.CurrentCulture.Name);
        Assert.NotEqual(".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
    }

    [Fact]
    public async Task NumbersAndDatesSurviveWhateverCultureWroteThem()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Данные");
        sheet["A1"] = 1234.5678;
        sheet["A2"] = -0.000123;
        sheet["A3"] = new DateTime(2026, 7, 26, 13, 45, 56, DateTimeKind.Unspecified);
        sheet["A4"] = TimeSpan.FromMinutes(93.5);
        sheet.Column("A").SetWidth(18.75);
        sheet.Row(1).SetHeight(22.5);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Worksheet read = reloaded.Sheets[0];
        Assert.Equal(1234.5678, read["A1"].GetNumber());
        Assert.Equal(-0.000123, read["A2"].GetNumber());
        Assert.Equal(new DateTime(2026, 7, 26, 13, 45, 56), read["A3"].GetDateTime());
        Assert.Equal(TimeSpan.FromMinutes(93.5), read["A4"].GetTimeSpan());
        Assert.Equal(18.75, read.ConfiguredColumns[1].Width);
        Assert.Equal(22.5, read.ConfiguredRows[1].Height);
    }

    /// <summary>
    /// The bytes themselves, not just the round trip: reading back through the same code would
    /// hide a separator that is wrong in both directions.
    /// </summary>
    [Fact]
    public async Task TheFileItselfSpellsNumbersWithAFullStop()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Data")["A1"] = 1234.5678;

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        using var archive = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        string sheet = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);

        Assert.Contains("1234.5678", sheet, StringComparison.Ordinal);
        Assert.DoesNotContain("1234,5678", sheet, StringComparison.Ordinal);
    }
}
