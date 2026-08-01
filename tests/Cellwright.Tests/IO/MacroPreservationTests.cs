using System.IO.Compression;
using System.Security.Cryptography;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// A macro-enabled workbook keeps its macros. The claim was in the README long before anything
/// checked it: the fidelity corpus only ever looked at <c>.xlsx</c>, so the one part with the
/// most to lose was never round-tripped at all.
/// </summary>
public sealed class MacroPreservationTests
{
    private static string? FindMacroWorkbook() => SampleCorpus.Files(int.MaxValue)
        .FirstOrDefault(static file => file.EndsWith("Control.xlsm", StringComparison.OrdinalIgnoreCase));

    [Fact]
    public async Task EditingACellLeavesTheVbaProjectByteForByte()
    {
        string? path = FindMacroWorkbook();
        Assert.SkipWhen(path is null, "No macro-enabled workbook is in the corpus. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        byte[] original = ReadPart(File.ReadAllBytes(path!), "xl/vbaProject.bin");

        Workbook book = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        book.Sheets[0]["ZZ100"] = "edited";

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        byte[] written = ReadPart(saved.ToArray(), "xl/vbaProject.bin");
        Assert.Equal(SHA256.HashData(original), SHA256.HashData(written));
    }

    /// <summary>
    /// The macro-enabled content type is what makes Excel offer to run the code. Writing the
    /// ordinary one back would leave the project in the package and unreachable.
    /// </summary>
    [Fact]
    public async Task TheWorkbookStaysMacroEnabled()
    {
        string? path = FindMacroWorkbook();
        Assert.SkipWhen(path is null, "No macro-enabled workbook is in the corpus. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        Workbook book = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string contentTypes = System.Text.Encoding.UTF8.GetString(ReadPart(saved.ToArray(), "[Content_Types].xml"));
        Assert.Contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml", contentTypes, StringComparison.Ordinal);
        Assert.Contains("application/vnd.ms-office.vbaProject", contentTypes, StringComparison.Ordinal);
    }

    private static byte[] ReadPart(byte[] package, string entryName)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        ZipArchiveEntry entry = Assert.IsType<ZipArchiveEntry>(archive.GetEntry(entryName));
        using var content = new MemoryStream();
        using (Stream source = entry.Open())
            source.CopyTo(content);
        return content.ToArray();
    }
}
