using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Opening the workbooks whose "password to open" has to be undone before a single record can
/// be read: the XOR obfuscation of the earliest files, and the two RC4 schemes after it.
/// </summary>
public sealed class EncryptionTests
{
    private static string? Find(string name) =>
        XlsCorpus.Files.FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    [Theory]
    [InlineData("Issue242_StdRc4PwdPassword.xls", "password")]   // RC4 CryptoAPI, SHA-1.
    [InlineData("Issue242_XorPwdPassword.xls", "password")]      // XOR obfuscation.
    [InlineData("Issue477_CryptoKeylength40.xls", "password")]   // RC4 CryptoAPI with a 40-bit key.
    public async Task AnEncryptedWorkbookOpensWithItsPassword(string name, string password)
    {
        string? path = Find(name);
        Assert.SkipWhen(path is null, $"{name} is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var warnings = new List<WorkbookWarning>();
        Workbook workbook = await XlsWorkbook.LoadAsync(
            path!,
            new XlsLoadOptions { Password = password, OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        Assert.NotEmpty(workbook.Sheets);
        Assert.Contains(warnings, static warning => warning.Code == "EncryptedWorkbook");

        // A file that decrypted to nonsense reads as an empty sheet, so the content is the
        // check that matters, not the absence of an exception.
        Worksheet sheet = workbook.Sheets[0];
        Assert.False(sheet.Cell(1, 1).Value.IsBlank, "The first cell is empty; decryption produced nothing.");
    }

    [Fact]
    public async Task TheContentOfAnEncryptedWorkbookIsWhatItSays()
    {
        string? path = Find("Issue242_StdRc4PwdPassword.xls");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(
            path!, new XlsLoadOptions { Password = "password" }, TestContext.Current.CancellationToken);

        Assert.Equal("Password: password", workbook.Sheets[0].Cell(1, 1).Value.GetText());
    }

    /// <summary>
    /// The corpus has no workbook using plain RC4 — the scheme Excel 97 used before the
    /// CryptoAPI one — and no library within reach writes one, so the fixture is built from
    /// the specification. The two RC4 schemes derive their keys differently, and a reader that
    /// muddles them decrypts to noise.
    /// </summary>
    [Fact]
    public async Task AWorkbookEncryptedWithPlainRc4Opens()
    {
        Workbook original = Workbook.Create();
        Worksheet source = original.Sheets.Add("Secret");
        source["A1"] = CellValue.Text("classified");
        source["A2"] = CellValue.Number(42.5);
        for (int row = 3; row <= 400; row++)
            source[row, 1] = CellValue.Text($"row {row} of a file long enough to span several cipher blocks");

        var warnings = new List<WorkbookWarning>();
        using var encrypted = new MemoryStream(await EncryptAsync(original, "secret"));
        Workbook workbook = await XlsWorkbook.LoadAsync(
            encrypted,
            new XlsLoadOptions { Password = "secret", LeaveOpen = true, OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        Worksheet read = workbook.Sheets[0];
        Assert.Equal("Secret", read.Name);
        Assert.Equal("classified", read["A1"].GetText());
        Assert.Equal(42.5, read["A2"].GetNumber());
        Assert.Equal("row 400 of a file long enough to span several cipher blocks", read[400, 1].GetText());
        Assert.Contains(warnings, static warning => warning.Code == "EncryptedWorkbook");
    }

    [Fact]
    public async Task PlainRc4RefusesTheWrongPassword()
    {
        Workbook original = Workbook.Create();
        original.Sheets.Add("Secret")["A1"] = CellValue.Text("classified");

        using var encrypted = new MemoryStream(await EncryptAsync(original, "secret"));
        await Assert.ThrowsAsync<XlsFormatException>(async () => await XlsWorkbook.LoadAsync(
            encrypted, new XlsLoadOptions { Password = "wrong", LeaveOpen = true }, TestContext.Current.CancellationToken));
    }

    /// <summary>Writes a workbook, then encrypts its stream and repackages it.</summary>
    private static async Task<byte[]> EncryptAsync(Workbook workbook, string password)
    {
        using var plain = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, plain, options: null, TestContext.Current.CancellationToken);

        Cellwright.Xls.Cfb.CompoundFile container = Cellwright.Xls.Cfb.CompoundFile.Parse(plain.ToArray());
        byte[] stream = LegacyRc4Encryptor.Encrypt(container.TryReadStream("Workbook")!, password);

        using var repacked = new MemoryStream();
        await Cellwright.Xls.Cfb.CompoundFileWriter.WriteAsync(
            repacked, "Workbook", stream, TestContext.Current.CancellationToken);
        return repacked.ToArray();
    }

    [Fact]
    public async Task TheWrongPasswordIsRefusedWithAnExplanation()
    {
        string? path = Find("Issue242_StdRc4PwdPassword.xls");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await XlsWorkbook.LoadAsync(
                path!, new XlsLoadOptions { Password = "not the password" }, TestContext.Current.CancellationToken));

        Assert.Contains("password", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AMissingPasswordSaysSoRatherThanFailingObscurely()
    {
        string? path = Find("Issue242_StdRc4PwdPassword.xls");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken));

        Assert.Contains("encrypted", failure.Message, StringComparison.OrdinalIgnoreCase);
    }
}
