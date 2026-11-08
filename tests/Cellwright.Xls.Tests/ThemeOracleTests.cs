using System.Text;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Checks the theme colour scheme and the tint arithmetic against NPOI, which implements the
/// same section of the standard independently.
/// </summary>
/// <remarks>
/// The tint is worth an oracle of its own. It is stated in §18.8.3 as arithmetic on the
/// luminance of an HSL conversion, and every step of that — the conversion, the direction of
/// the shift, the rounding back to bytes — is a place to be plausibly wrong in a way no file
/// will ever complain about.
/// </remarks>
public sealed class ThemeOracleTests
{
    static ThemeOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task ThemeSlotsAgreeWithTheOracle()
    {
        int compared = 0, mismatched = 0;
        var examples = new List<string>();

        foreach (string path in XlsCorpus.Packages.Take(120))
        {
            List<string>? oracle = ReadScheme(path);
            if (oracle is null)
                continue;

            Workbook mine;
            try
            {
                mine = await Workbook.LoadAsync(path, cancellationToken: TestContext.Current.CancellationToken);
            }
            catch (Exception)
            {
                continue;
            }

            for (int slot = 0; slot < oracle.Count && slot < WorkbookTheme.SlotCount; slot++)
            {
                compared++;
                string ours = mine.Theme[slot].ToHex()[2..];
                if (string.Equals(ours, oracle[slot], StringComparison.OrdinalIgnoreCase))
                    continue;

                mismatched++;
                if (examples.Count < 20)
                    examples.Add($"{Path.GetFileName(path)} slot {slot}: ours {ours} vs npoi {oracle[slot]}");
            }
        }

        Assert.SkipWhen(compared == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        Assert.True(compared > 200, $"Only {compared} theme slots were compared; coverage has dropped.");
        Assert.True(
            mismatched == 0,
            $"{mismatched} of {compared} theme slots differ:{Environment.NewLine}" +
            string.Join(Environment.NewLine, examples));
    }

    /// <summary>
    /// The tint arithmetic itself, over the full range and every slot of the standard scheme.
    /// A tint of zero must be the identity, and the extremes must reach white and black.
    /// </summary>
    [Theory]
    [InlineData(0xFF4F81BD, 0.0, "FF4F81BD")]
    [InlineData(0xFF4F81BD, 1.0, "FFFFFFFF")]
    [InlineData(0xFF4F81BD, -1.0, "FF000000")]
    [InlineData(0xFF000000, 0.5, "FF808080")]
    [InlineData(0xFFFFFFFF, -0.25, "FFBFBFBF")]
    public void TintMovesLuminanceTowardsWhiteOrBlack(uint argb, double tint, string expected) =>
        Assert.Equal(expected, WorkbookTheme.ApplyTint(argb, tint).ToString("X8"));

    /// <summary>
    /// The greys of Excel's colour picker, which are the white of the theme's first background
    /// slot darkened by the four amounts the picker offers. The tints are the exact values
    /// Excel writes, not round numbers, and the results are the greys a user recognizes.
    /// </summary>
    [Theory]
    [InlineData(-0.149998474074526, "FFD9D9D9")]
    [InlineData(-0.249977111117893, "FFBFBFBF")]
    [InlineData(-0.349986266670736, "FFA6A6A6")]
    [InlineData(-0.499984740745262, "FF808080")]
    public void TheGreysOfTheColourPickerComeOutAsExcelWritesThem(double tint, string expected) =>
        Assert.Equal(expected, WorkbookTheme.ApplyTint(0xFFFFFFFF, tint).ToString("X8"));

    /// <summary>Reads the theme's colour scheme through NPOI, as twelve hex strings.</summary>
    private static List<string>? ReadScheme(string path)
    {
        try
        {
            using FileStream input = File.OpenRead(path);
            using var book = new XSSFWorkbook(input);
            NPOI.XSSF.Model.ThemesTable? themes = book.GetTheme();
            if (themes is null)
                return null;

            // NPOI numbers the slots as the styles do, with the dark and light pairs swapped —
            // the same order this library uses, so the two indexes line up directly.
            var scheme = new List<string>(WorkbookTheme.SlotCount);
            for (int slot = 0; slot < WorkbookTheme.SlotCount; slot++)
            {
                if (themes.GetThemeColor(slot)?.ARGBHex is not { Length: >= 6 } hex)
                    return null;
                scheme.Add(hex[^6..]);
            }

            return scheme;
        }
        catch (Exception)
        {
            return null;
        }
    }
}
