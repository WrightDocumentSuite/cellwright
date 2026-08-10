using System.Text;
using Cellwright.Xls.Cfb;
using Cellwright.Xls.Crypto;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Opening an .xlsx locked with a password to open. Excel does not encrypt the zip: it puts the
/// whole package inside a compound file, so such a workbook arrives looking exactly like a
/// legacy one and every reader that trusts the first eight bytes gets it wrong.
/// </summary>
public sealed class EncryptedPackageTests
{
    private const string Password = "password";

    private static string? Find(string name) => XlsCorpus.Packages
        .FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    [Theory]
    [InlineData("standard_AES128_SHA1_ECB_pwd_password.xlsx")]
    [InlineData("standard_AES192_SHA1_ECB_pwd_password.xlsx")]
    [InlineData("standard_AES256_SHA1_ECB_pwd_password.xlsx")]
    [InlineData("agile_AES128_MD5_CBC_pwd_password.xlsx")]
    [InlineData("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [InlineData("agile_AES128_SHA384_CBC_pwd_password.xlsx")]
    [InlineData("agile_AES128_SHA512_CBC_pwd_password.xlsx")]
    [InlineData("agile_AES192_SHA512_CBC_pwd_password.xlsx")]
    [InlineData("agile_AES256_SHA512_CBC_pwd_password.xlsx")]
    [InlineData("agile_DES_MD5_CBC_pwd_password.xlsx")]
    [InlineData("agile_DESede_SHA384_CBC_pwd_password.xlsx")]
    [InlineData("agile_RC2_SHA1_CBC_pwd_password.xlsx")]
    public async Task AnEncryptedPackageOpensWithItsPassword(string name)
    {
        string? path = Find(name);
        Assert.SkipWhen(path is null, $"{name} is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var warnings = new List<WorkbookWarning>();
        Workbook workbook = await Spreadsheet.LoadAsync(
            path!,
            new XlsLoadOptions { Password = Password, OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        // Decrypting to noise produces an unreadable zip rather than a wrong value, but the
        // content is still what proves the key was right rather than merely plausible.
        Assert.Equal("Password: password", workbook.Sheets[0].Cell(1, 1).Value.GetText());
        Assert.Contains(warnings, static warning => warning.Code == "EncryptedWorkbook");

        // The agile files carry an HMAC over the encrypted package. Checking it is only worth
        // anything if it holds on a file nobody has touched, which is what this asserts.
        Assert.DoesNotContain(warnings, static warning => warning.Code == "PackageIntegrityFailed");
    }

    /// <summary>
    /// A package altered after it was encrypted still decrypts — the cipher has no opinion —
    /// but the integrity value stored beside it no longer matches, and that is worth saying.
    /// </summary>
    [Fact]
    public async Task AnAlteredPackageDecryptsUnderProtest()
    {
        string? path = Find("agile_AES256_SHA512_CBC_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        byte[] bytes = await File.ReadAllBytesAsync(path!, TestContext.Current.CancellationToken);
        Assert.True(Corrupt(bytes), "The encrypted package was not found in the file.");

        var warnings = new List<WorkbookWarning>();
        using var stream = new MemoryStream(bytes, writable: false);
        try
        {
            await Spreadsheet.LoadAsync(
                stream,
                new XlsLoadOptions { Password = Password, OnWarning = warnings.Add },
                TestContext.Current.CancellationToken);
        }
        catch (Exception exception) when (exception is XlsxFormatException or InvalidDataException)
        {
            // Flipping a byte of ciphertext ruins a block of the zip; the warning comes first.
        }

        Assert.Contains(warnings, static warning => warning.Code == "PackageIntegrityFailed");
    }

    /// <summary>
    /// Flips one byte of ciphertext. The package is a stream inside a compound file, so the
    /// byte to alter is found by looking for a window of the stream's own content in the raw
    /// file — guessing an offset would as likely land in the container's bookkeeping.
    /// </summary>
    private static bool Corrupt(byte[] bytes)
    {
        byte[]? package = CompoundFile.Parse(bytes).TryReadStream("EncryptedPackage");
        if (package is null || package.Length < 128)
            return false;

        int at = bytes.AsSpan().IndexOf(package.AsSpan(64, 32));
        if (at < 0)
            return false;

        bytes[at] ^= 0xFF;
        return true;
    }

    /// <summary>
    /// Agile encryption names its algorithms in the file rather than fixing them, and the
    /// format allows more of them than any implementation carries. One it does not know has to
    /// be refused by name rather than decrypted into rubbish.
    /// </summary>
    [Fact]
    public void ACipherWeDoNotCarryIsNamedRatherThanGuessedAt()
    {
        XlsFormatException failure = Assert.Throws<XlsFormatException>(() =>
            Read(cipher: "DESX", keyBits: 128, blockSize: 8));

        Assert.Contains("DESX", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A key length the named cipher does not have means the descriptor and the data disagree,
    /// and decrypting anyway would produce plausible-looking noise.
    /// </summary>
    [Fact]
    public void AKeyLengthTheCipherDoesNotHaveIsRefused()
    {
        XlsFormatException failure = Assert.Throws<XlsFormatException>(() =>
            Read(cipher: "DES", keyBits: 256, blockSize: 8));

        Assert.Contains("DES", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>A block size that contradicts the cipher is the same kind of disagreement.</summary>
    [Fact]
    public void ABlockSizeThatContradictsTheCipherIsRefused()
    {
        XlsFormatException failure = Assert.Throws<XlsFormatException>(() =>
            Read(cipher: "AES", keyBits: 128, blockSize: 8));

        Assert.Contains("16", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOversizedCryptoMetadataAttributeIsRefusedBeforeBase64BufferAllocation()
    {
        string descriptor = Descriptor("AES", keyBits: 128, blockSize: 16).Replace(
            "encryptedKeyValue=\"AAAAAAAAAAAAAAAAAAAAAA==\"",
            $"encryptedKeyValue=\"{new string('A', (4 * 1024) + 1)}\"",
            StringComparison.Ordinal);

        XlsFormatException failure = Assert.Throws<XlsFormatException>(() =>
            AgileEncryption.Decrypt(
                Encoding.UTF8.GetBytes(descriptor),
                new byte[8],
                [Password],
                cancellationToken: TestContext.Current.CancellationToken));

        Assert.Contains("encryptedKeyValue", failure.Message, StringComparison.Ordinal);
        Assert.Contains("4096", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnOversizedEncryptionDescriptorIsRefusedBeforeItsCfbStreamIsMaterialized()
    {
        byte[] oversizedInfo = new byte[(1 * 1024 * 1024) + 9];
        using var encrypted = new MemoryStream();
        await CompoundFileWriter.WriteAsync(
            encrypted,
            [
                new CompoundFileWriter.NamedStream("EncryptionInfo", oversizedInfo),
                new CompoundFileWriter.NamedStream("EncryptedPackage", new byte[8]),
            ],
            TestContext.Current.CancellationToken);
        encrypted.Position = 0;

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await Spreadsheet.LoadAsync(
                encrypted,
                new XlsLoadOptions { LeaveOpen = true, Password = Password },
                TestContext.Current.CancellationToken));

        Assert.Contains("EncryptionInfo", failure.Message, StringComparison.Ordinal);
        Assert.Contains("1048584", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// Runs the descriptor through the reader. The package handed over is empty because these
    /// cases all fail while reading the descriptor, before any byte of it is touched.
    /// </summary>
    private static void Read(string cipher, int keyBits, int blockSize) =>
        AgileEncryption.Decrypt(
            Encoding.UTF8.GetBytes(Descriptor(cipher, keyBits, blockSize)),
            new byte[8],
            [Password]);

    /// <summary>The smallest descriptor that parses, with one attribute under test.</summary>
    private static string Descriptor(string cipher, int keyBits, int blockSize) =>
        $"""
         <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
         <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
           <keyData saltSize="16" blockSize="{blockSize}" keyBits="{keyBits}" hashSize="20"
                    cipherAlgorithm="{cipher}" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                    saltValue="AAAAAAAAAAAAAAAAAAAAAA==" />
           <keyEncryptors>
             <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
               <encryptedKey spinCount="100000" saltSize="16" blockSize="{blockSize}" keyBits="{keyBits}"
                             hashSize="20" cipherAlgorithm="{cipher}" cipherChaining="ChainingModeCBC"
                             hashAlgorithm="SHA1" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="
                             encryptedVerifierHashInput="AAAAAAAAAAAAAAAAAAAAAA=="
                             encryptedVerifierHashValue="AAAAAAAAAAAAAAAAAAAAAA=="
                             encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAA==" />
             </keyEncryptor>
           </keyEncryptors>
         </encryption>
         """;

    [Fact]
    public async Task TheWrongPasswordIsRefusedWithAnExplanation()
    {
        string? path = Find("agile_AES256_SHA512_CBC_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await Spreadsheet.LoadAsync(
                path!, new XlsLoadOptions { Password = "not the password" }, TestContext.Current.CancellationToken));

        Assert.Contains("password", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AMissingPasswordSaysWhereToPutOne()
    {
        string? path = Find("standard_AES128_SHA1_ECB_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await Spreadsheet.LoadAsync(path!, options: null, TestContext.Current.CancellationToken));

        Assert.Contains("XlsLoadOptions.Password", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The legacy reader is the one such a file reaches by mistake, since the container is the
    /// same. It has to say what the file actually is instead of reporting a missing stream.
    /// </summary>
    [Fact]
    public async Task TheLegacyReaderRefusesAnEncryptedPackageByName()
    {
        string? path = Find("agile_AES128_SHA1_CBC_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await XlsWorkbook.LoadAsync(
                path!, new XlsLoadOptions { Password = Password }, TestContext.Current.CancellationToken));

        Assert.Contains("Spreadsheet.LoadAsync", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>An encrypted package still preserves what the model does not represent.</summary>
    [Fact]
    public async Task ADecryptedPackageSavesBackAsAnOrdinaryWorkbook()
    {
        string? path = Find("agile_AES256_SHA512_CBC_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await Spreadsheet.LoadAsync(
            path!, new XlsLoadOptions { Password = Password }, TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await workbook.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("Password: password", reloaded.Sheets[0].Cell(1, 1).Value.GetText());
    }
}
