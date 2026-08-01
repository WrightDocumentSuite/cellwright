using System.Text;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Named fixtures from the corpus, each read for the one thing it was made to exercise. The
/// corpus-wide oracle says the reader agrees with another implementation; these say what it
/// agrees about, and they fail with a legible message when one corner breaks.
/// </summary>
public sealed class XlsbReadingTests
{
    static XlsbReadingTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>Loads a fixture, skipping the test when the corpus is not on this machine.</summary>
    private static async Task<Workbook> OpenAsync(string name, string? password = null)
    {
        string? path = XlsbCorpus.Files.FirstOrDefault(
            file => string.Equals(Path.GetFileName(file), name, StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, $"{name} is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var options = password is null ? null : new XlsLoadOptions { Password = password };
        return await Spreadsheet.LoadAsync(path!, options, TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task TheSheetsArriveInTabOrderWithTheirOwnNames()
    {
        Workbook workbook = await OpenAsync("MultiSheet.xlsb");

        Assert.Equal(["Sheet2", "Sheet1", "Sheet3"], workbook.Sheets.Select(static sheet => sheet.Name));
        Assert.Equal(1, workbook.Sheets[0][1, 1].GetNumber());
        Assert.Equal(2, workbook.Sheets[1][1, 1].GetNumber());
        Assert.Equal(3, workbook.Sheets[2][1, 1].GetNumber());
    }

    /// <summary>
    /// Every sheet of this file is written with four bytes ahead of the tab identifier that the
    /// documented layout does not have. Read at the documented offset the relationship id comes
    /// out as a single null character, which matches nothing, and the workbook loads with no
    /// sheets at all rather than failing.
    /// </summary>
    [Fact]
    public async Task ASheetListWithAnExtraFieldStillNamesItsParts()
    {
        Workbook workbook = await OpenAsync("Simple.xlsb");

        Assert.Equal(["Sheet1", "Sheet2", "Sheet3"], workbook.Sheets.Select(static sheet => sheet.Name));
        Assert.StartsWith(
            "This is an example spreadsheet",
            workbook.Sheets[0][1, 1].GetText(),
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheEpochComesFromTheWorkbookRatherThanTheReader()
    {
        Workbook nineteenHundred = await OpenAsync("roo_1900_base.xlsb");
        Workbook nineteenOhFour = await OpenAsync("roo_1904_base.xlsb");

        Assert.Equal(ExcelDateSystem.Date1900, nineteenHundred.DateSystem);
        Assert.Equal(ExcelDateSystem.Date1904, nineteenOhFour.DateSystem);

        // Both files hold the same dates; only the serial numbers under them differ.
        Assert.Equal(new DateTime(2009, 6, 15), nineteenHundred.Sheets[0][1, 1].GetDateTime());
        Assert.Equal(new DateTime(2009, 6, 15), nineteenOhFour.Sheets[0][1, 1].GetDateTime());
    }

    /// <summary>
    /// Text formatted in pieces is written into the cell itself rather than into the shared
    /// table, in a record the older format has no equivalent of.
    /// </summary>
    [Fact]
    public async Task RichTextArrivesAsPlainText()
    {
        Workbook workbook = await OpenAsync("Issue250_Richtext.xlsb");

        Assert.StartsWith("Lorem ipsum dolor sit amet", workbook.Sheets[0][1, 1].GetText(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnErrorKeepsItsOwnKind()
    {
        Workbook workbook = await OpenAsync("Issue329_Error.xlsb");
        Worksheet sheet = workbook.Sheets[0];

        Assert.Equal(ExcelError.DivisionByZero, sheet[1, 1].GetError());
        Assert.Equal(ExcelError.NotAvailable, sheet[1, 2].GetError());
        Assert.Equal(ExcelError.Value, sheet[1, 3].GetError());
    }

    [Fact]
    public async Task AFormulaKeepsBothItsTextAndItsLastResult()
    {
        Workbook workbook = await OpenAsync("BoolFormula.xlsb");
        Worksheet sheet = workbook.Sheets[0];

        Assert.Equal("(1=1)", sheet.Cell(1, 1).Formula);
        Assert.True(sheet[1, 1].GetBoolean());
    }

    /// <summary>
    /// Numbers survive as doubles rather than being rounded through a decimal on the way in,
    /// including the exponents at either end of the range.
    /// </summary>
    [Fact]
    public async Task NumbersKeepEveryDigitTheyWereStoredWith()
    {
        Workbook workbook = await OpenAsync("DoublePrecision.xlsb");
        Worksheet sheet = workbook.Sheets[0];

        // The file itself holds pi to fifteen digits, which is what Excel stores; the point is
        // that the reader hands back that double and not a rounded rendering of it.
        Assert.Equal(3.14159265358979, sheet[3, 2].GetNumber());
        Assert.Equal(3.14159265358979e-300, sheet[5, 2].GetNumber());
        Assert.Equal(3.141592653589793e+300, sheet[7, 2].GetNumber());
        Assert.Equal(-314159265358979d, sheet[10, 2].GetNumber());
    }

    [Fact]
    public async Task MergedRangesComeThrough()
    {
        Workbook workbook = await OpenAsync("MergedCell.xlsb");

        Assert.Equal(4, workbook.Sheets[0].MergedRanges.Count);
        Assert.Equal("Merge Cell 1", workbook.Sheets[0][2, 1].GetText());
    }

    /// <summary>
    /// An encrypted binary workbook is a compound file wrapping the package, exactly as an
    /// encrypted xlsx is, so it costs the reader nothing beyond sending the decrypted bytes
    /// back through the same fork in the road. That holds for every cipher, not just the AES
    /// that Office writes by default.
    /// </summary>
    [Theory]
    [InlineData("standard_AES256_SHA1_ECB_pwd_password.xlsb")]
    [InlineData("agile_AES128_MD5_CBC_pwd_password.xlsb")]
    [InlineData("agile_DES_MD5_CBC_pwd_password.xlsb")]
    [InlineData("agile_DESede_SHA384_CBC_pwd_password.xlsb")]
    [InlineData("agile_RC2_SHA1_CBC_pwd_password.xlsb")]
    public async Task AnEncryptedWorkbookOpensWithItsPassword(string name)
    {
        Workbook workbook = await OpenAsync(name, password: "password");

        Assert.Equal("Sheet1", workbook.Sheets[0].Name);
        Assert.NotNull(workbook.Sheets[0].UsedRange);
    }

    /// <summary>A sheet of ten thousand rows exercises the row header, which cells hang off.</summary>
    [Fact]
    public async Task ALongSheetKeepsItsRowsInOrder()
    {
        Workbook workbook = await OpenAsync("10x10000.xlsb");
        Worksheet sheet = workbook.Sheets[0];

        Assert.Equal(10_000, sheet.UsedRange?.End.Row);
        Assert.Equal(10, sheet.UsedRange?.End.Column);
    }

    [Fact]
    public async Task HiddenAndCollapsedRowsAndColumnsAreRecorded()
    {
        Workbook workbook = await OpenAsync("CollapsedHide.xlsb");
        Worksheet sheet = workbook.Sheets[0];

        Assert.True(sheet.Column(1).Hidden);
        Assert.True(sheet.Row(4).Hidden);
    }
}
