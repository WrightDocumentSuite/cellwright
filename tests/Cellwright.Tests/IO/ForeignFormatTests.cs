using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// Files that are plainly not an .xlsx package have to say so. Every one of these used to reach
/// the reader and fail somewhere deeper — a missing part, an XML parse error, or a message
/// naming the wrong format entirely — leaving whoever hit it to guess.
/// </summary>
public sealed class ForeignFormatTests
{
    [Fact]
    public async Task ABinaryWorkbookIsNamedRatherThanParsedAsXml()
    {
        string? path = SampleCorpus.Find("10x10.xlsb");
        Assert.SkipWhen(path is null, "No .xlsb is in the corpus. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        XlsxFormatException failure = await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await Workbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken));

        Assert.Contains(".xlsb", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnOpenDocumentSpreadsheetIsNamedRatherThanReportedAsAMissingPart()
    {
        string? path = SampleCorpus.Find("MySheets.ods");
        Assert.SkipWhen(path is null, "No .ods is in the corpus. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        XlsxFormatException failure = await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await Workbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken));

        Assert.Contains("OpenDocument", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnOrdinaryZipIsNotMistakenForAPackage()
    {
        using var zip = new MemoryStream();
        using (var archive = new ZipArchive(zip, ZipArchiveMode.Create, leaveOpen: true))
        {
            await using Stream entry = archive.CreateEntry("notes.txt").Open();
            await entry.WriteAsync(Encoding.UTF8.GetBytes("nothing to see"), TestContext.Current.CancellationToken);
        }

        zip.Position = 0;
        XlsxFormatException failure = await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await Workbook.LoadAsync(zip, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken));

        Assert.Contains("[Content_Types].xml", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A compound file is either a legacy workbook or a package encrypted with a password to
    /// open, and the core cannot tell which without a container parser. Naming only the first
    /// used to send everyone holding an encrypted .xlsx down the wrong path entirely.
    /// </summary>
    [Fact]
    public async Task ACompoundFileNamesBothThingsItCouldBe()
    {
        string? path = SampleCorpus.Find("agile_AES128_SHA1_CBC_pwd_password.xlsx");
        Assert.SkipWhen(path is null, "No encrypted package is in the corpus. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        XlsxFormatException failure = await Assert.ThrowsAsync<XlsxFormatException>(async () =>
            await Workbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken));

        Assert.Contains("encrypted", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Cellwright.Xls", failure.Message, StringComparison.Ordinal);
    }
}
