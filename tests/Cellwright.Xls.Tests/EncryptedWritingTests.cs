using System.Text;
using Cellwright.Xls.Cfb;
using Cellwright.Xls.Crypto;
using ExcelDataReader;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Writing a package behind a password to open.
/// </summary>
/// <remarks>
/// Encryption is the one area where a round trip through this library proves almost nothing: a
/// reader and a writer that share a misunderstanding agree with each other perfectly, and the
/// file they produce is one nothing else can open. So the checks that matter here are the two
/// foreign ones — NPOI and ExcelDataReader, each with its own reading of MS-OFFCRYPTO — and the
/// round trip is only there to say what the content should be.
/// </remarks>
public sealed class EncryptedWritingTests
{
    private const string Password = "correct horse battery staple";

    static EncryptedWritingTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task AnEncryptedWorkbookOpensWithItsPassword()
    {
        byte[] file = await WriteAsync(Sample());

        var warnings = new List<WorkbookWarning>();
        using var stream = new MemoryStream(file, writable: false);
        Workbook reloaded = await Spreadsheet.LoadAsync(
            stream, new XlsLoadOptions { Password = Password, OnWarning = warnings.Add }, TestContext.Current.CancellationToken);

        Worksheet sheet = reloaded.Sheets[0];
        Assert.Equal("hidden in plain sight", sheet.Cell(1, 1).Value.GetText());
        Assert.Equal(1234.5, sheet.Cell(2, 1).Value.GetNumber());
        Assert.Equal("SUM(A2:A2)", sheet.Cell(3, 1).Formula);
        Assert.Contains(warnings, static warning => warning.Code == "EncryptedWorkbook");

        // The integrity value is written over the finished ciphertext, so a reader that checks
        // it has to find it holds. Silence here is the check.
        Assert.DoesNotContain(warnings, static warning => warning.Code == "PackageIntegrityFailed");
    }

    /// <summary>
    /// NPOI reads the file through its own implementation of the same specification. This is
    /// the assertion the whole feature rests on: agreement between our writer and our reader
    /// would be equally perfect if both were wrong.
    /// </summary>
    [Fact]
    public async Task NpoiOpensWhatWeWrote()
    {
        byte[] file = await WriteAsync(Sample());

        using var stream = new MemoryStream(file, writable: false);
        var container = new POIFSFileSystem(stream);
        Decryptor decryptor = Decryptor.GetInstance(new EncryptionInfo(container));

        Assert.True(decryptor.VerifyPassword(Password), "NPOI does not accept the password we wrote.");
        Assert.False(decryptor.VerifyPassword("something else"), "NPOI accepts a password that is not the one.");

        Assert.True(decryptor.VerifyPassword(Password));
        using var workbook = new XSSFWorkbook(decryptor.GetDataStream(container));
        Assert.Equal("hidden in plain sight", workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>A second foreign reader, on a different code path from NPOI's.</summary>
    [Fact]
    public async Task ExcelDataReaderOpensWhatWeWrote()
    {
        byte[] file = await WriteAsync(Sample());

        using var stream = new MemoryStream(file, writable: false);
        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(
            stream, new ExcelReaderConfiguration { Password = Password });

        Assert.Equal("Secrets", reader.Name);
        Assert.True(reader.Read());
        Assert.Equal("hidden in plain sight", reader.GetValue(0));
    }

    /// <summary>
    /// A package larger than one segment, so that the per-segment vectors are exercised rather
    /// than a single block. Getting the segment numbering wrong shows up nowhere else: the
    /// first 4096 bytes decrypt correctly either way.
    /// </summary>
    [Fact]
    public async Task AWorkbookOfManySegmentsSurvives()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Long");
        for (int row = 1; row <= 4000; row++)
            sheet[row, 1] = $"row {row} of a package that runs well past one segment";

        byte[] file = await WriteAsync(book);
        Assert.True(file.Length > 3 * 4096, $"The sample is only {file.Length} bytes; it does not span segments.");

        using var stream = new MemoryStream(file, writable: false);
        Workbook reloaded = await Spreadsheet.LoadAsync(
            stream, new XlsLoadOptions { Password = Password }, TestContext.Current.CancellationToken);

        Assert.Equal("row 1 of a package that runs well past one segment", reloaded.Sheets[0].Cell(1, 1).Value.GetText());
        Assert.Equal("row 4000 of a package that runs well past one segment", reloaded.Sheets[0].Cell(4000, 1).Value.GetText());
    }

    /// <summary>
    /// The parts the model does not understand still travel: encrypting wraps the very package
    /// the ordinary saver produces, so nothing is traded away for the password. A workbook with
    /// pivot tables makes the point, since those are precisely what the model does not model.
    /// </summary>
    [Fact]
    public async Task EncryptingCostsNothingTheOrdinarySaveKeeps()
    {
        string? path = XlsCorpus.Packages.FirstOrDefault(
            static p => Path.GetFileName(p).Equals("PivotTables.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "The sample is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook original = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        byte[] file = await WriteAsync(original);

        using var stream = new MemoryStream(file, writable: false);
        Workbook reloaded = await Spreadsheet.LoadAsync(
            stream, new XlsLoadOptions { Password = Password }, TestContext.Current.CancellationToken);

        Assert.Equal(original.Sheets.Count, reloaded.Sheets.Count);
        for (int i = 0; i < original.Sheets.Count; i++)
        {
            Assert.Equal(original.Sheets[i].Name, reloaded.Sheets[i].Name);
            Assert.Equal(original.Sheets[i].UsedRange?.ToString(), reloaded.Sheets[i].UsedRange?.ToString());
        }

        // The pivot parts are inside the encrypted package, which is the whole claim: what the
        // password protects is an ordinary xlsx, not a reduced one.
        Assert.Contains(EntryNames(file, Password), static name => name.Contains("pivot", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Part names of the package inside, read back through the reader under test.</summary>
    private static List<string> EntryNames(byte[] file, string password)
    {
        byte[] package = OoxmlEncryption.Decrypt(CompoundFile.Parse(file), password);
        using var stream = new MemoryStream(package, writable: false);
        using var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
        return [.. zip.Entries.Select(static entry => entry.FullName)];
    }

    [Fact]
    public async Task TheWrongPasswordIsRefused()
    {
        byte[] file = await WriteAsync(Sample());

        using var stream = new MemoryStream(file, writable: false);
        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await Spreadsheet.LoadAsync(
                stream, new XlsLoadOptions { Password = "not the password" }, TestContext.Current.CancellationToken));

        Assert.Contains("password", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NoPasswordAtAllSaysWhereToPutOne()
    {
        byte[] file = await WriteAsync(Sample());

        using var stream = new MemoryStream(file, writable: false);
        XlsFormatException failure = await Assert.ThrowsAsync<XlsFormatException>(async () =>
            await Spreadsheet.LoadAsync(stream, options: null, TestContext.Current.CancellationToken));

        Assert.Contains("XlsLoadOptions.Password", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// What we write is a compound file, not a zip, so the core reader cannot open it and has
    /// to say why rather than complain about a corrupt package.
    /// </summary>
    [Fact]
    public async Task TheFileIsACompoundFileAndTheCoreReaderSaysSo()
    {
        byte[] file = await WriteAsync(Sample());

        Assert.Equal("D0CF11E0A1B11AE1", Convert.ToHexString(file.AsSpan(0, 8)));

        using var stream = new MemoryStream(file, writable: false);
        XlsxFormatException failure = await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await Workbook.LoadAsync(stream, cancellationToken: TestContext.Current.CancellationToken));

        Assert.Contains("Cellwright.Xls", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A file altered after we wrote it fails the integrity check we wrote into it. Without
    /// this the HMAC could be a constant nobody ever compares against.
    /// </summary>
    [Fact]
    public async Task OurOwnIntegrityValueCatchesAnAlteredPackage()
    {
        byte[] file = await WriteAsync(Sample());
        Assert.True(Corrupt(file), "The encrypted package was not found in the file we wrote.");

        var warnings = new List<WorkbookWarning>();
        using var stream = new MemoryStream(file, writable: false);
        try
        {
            await Spreadsheet.LoadAsync(
                stream,
                new XlsLoadOptions { Password = Password, OnWarning = warnings.Add },
                TestContext.Current.CancellationToken);
        }
        catch (Exception exception) when (exception is XlsxFormatException or InvalidDataException)
        {
            // Flipping ciphertext ruins a block of the zip inside; the warning comes first.
        }

        Assert.Contains(warnings, static warning => warning.Code == "PackageIntegrityFailed");
    }

    /// <summary>
    /// Two files written from the same workbook under the same password differ, because the
    /// salts and the key are drawn fresh each time. Identical output would mean the randomness
    /// is not there, which is the kind of defect that never shows up as a failure to open.
    /// </summary>
    [Fact]
    public async Task TwoWritesOfOneWorkbookAreNotTheSameBytes()
    {
        byte[] first = await WriteAsync(Sample());
        byte[] second = await WriteAsync(Sample());

        Assert.NotEqual(first, second);
    }

    [Fact]
    public async Task AnEmptyPasswordIsRefusedRatherThanWrittenAsNoPassword()
    {
        using var stream = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(async () =>
            await Sample().SaveEncryptedAsync(stream, "", null, TestContext.Current.CancellationToken));
    }

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Secrets");
        sheet["A1"] = "hidden in plain sight";
        sheet["A2"] = 1234.5;
        sheet.Cell("A3").SetFormula("SUM(A2:A2)");
        return book;
    }

    private static async Task<byte[]> WriteAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveEncryptedAsync(stream, Password, null, TestContext.Current.CancellationToken);
        return stream.ToArray();
    }

    /// <summary>Flips one byte of the ciphertext, found by its content rather than by offset.</summary>
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
}
