using System.Text;
using NPOI.HSSF.UserModel;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The title, author and the rest, which a compound file keeps in property set streams beside
/// the workbook rather than inside it.
/// </summary>
/// <remarks>
/// Checked against NPOI, which reads the same streams through its own implementation. The value
/// of the oracle here is that these streams are self-describing only up to a point: the strings
/// are in whatever code page the set declares, and a reader that assumes one produces text that
/// looks right in English and wrong in everything else.
/// </remarks>
public sealed class DocumentPropertyTests
{
    static DocumentPropertyTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task DocumentPropertiesAgreeWithTheOracle()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int compared = 0;
        var differences = new List<string>();

        foreach (string path in XlsCorpus.Files)
        {
            Workbook workbook;
            try
            {
                workbook = await XlsWorkbook.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            if (ReadWithOracle(path) is not { } oracle)
                continue;

            Check(path, "Title", oracle.Title, workbook.Properties.Title, ref compared, differences);
            Check(path, "Subject", oracle.Subject, workbook.Properties.Subject, ref compared, differences);
            Check(path, "Author", oracle.Author, workbook.Properties.Author, ref compared, differences);
            Check(path, "Keywords", oracle.Keywords, workbook.Properties.Keywords, ref compared, differences);
            Check(path, "LastAuthor", oracle.LastAuthor, workbook.Properties.LastModifiedBy, ref compared, differences);
            Check(path, "Application", oracle.Application, workbook.Properties.Application, ref compared, differences);
            Check(path, "Company", oracle.Company, workbook.Properties.Company, ref compared, differences);
            Check(path, "Manager", oracle.Manager, workbook.Properties.Manager, ref compared, differences);
            Check(path, "Category", oracle.Category, workbook.Properties.Category, ref compared, differences);
        }

        Assert.True(compared > 200, $"Only {compared} properties were comparable; coverage has dropped.");
        Assert.True(
            differences.Count == 0,
            $"{differences.Count} of {compared} properties differ from the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, differences.Take(25)));
    }

    /// <summary>
    /// A property survives being written and read back. The legacy writer has to create streams
    /// beside the workbook to carry it, which is a shape the container writer did not have.
    /// </summary>
    [Fact]
    public async Task PropertiesSurviveTheRoundTrip()
    {
        Workbook original = Workbook.Create();
        original.Sheets.Add("Sheet1")["A1"] = CellValue.Number(1);
        original.Properties.Title = "Quarterly figures";
        original.Properties.Author = "Отдел финансов";
        original.Properties.Company = "Example Ltd";
        original.Properties.Keywords = "budget, 2026";
        original.Properties.Created = new DateTime(2026, 1, 5, 9, 30, 0, DateTimeKind.Utc);

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        byte[] bytes = buffer.ToArray();

        using var forUs = new MemoryStream(bytes, writable: false);
        Workbook reread = await XlsWorkbook.LoadAsync(forUs, options: null, TestContext.Current.CancellationToken);

        Assert.Equal("Quarterly figures", reread.Properties.Title);
        Assert.Equal("Отдел финансов", reread.Properties.Author);
        Assert.Equal("Example Ltd", reread.Properties.Company);
        Assert.Equal("budget, 2026", reread.Properties.Keywords);
        Assert.Equal(new DateTime(2026, 1, 5, 9, 30, 0, DateTimeKind.Utc), reread.Properties.Created);

        // And the other reader agrees, which is what says the streams are shaped as Excel's are
        // rather than merely as this library's own reader expects.
        using var forOracle = new MemoryStream(bytes, writable: false);
        var oracle = new HSSFWorkbook(forOracle);
        Assert.Equal("Quarterly figures", oracle.SummaryInformation.Title);
        Assert.Equal("Отдел финансов", oracle.SummaryInformation.Author);
        Assert.Equal("Example Ltd", oracle.DocumentSummaryInformation.Company);
    }

    private static void Check(
        string path, string name, string? expected, string? actual, ref int compared, List<string> differences)
    {
        if (string.IsNullOrEmpty(expected))
            return;

        compared++;
        if (string.Equals(expected.Trim(), actual?.Trim(), StringComparison.Ordinal))
            return;

        differences.Add($"{Path.GetFileName(path)} {name}: oracle '{expected}' vs '{actual}'");
    }

    private readonly record struct OracleProperties(
        string? Title, string? Subject, string? Author, string? Keywords,
        string? LastAuthor, string? Application, string? Company, string? Manager, string? Category);

    private static OracleProperties? ReadWithOracle(string path)
    {
        try
        {
            using FileStream stream = File.OpenRead(path);
            var workbook = new HSSFWorkbook(stream);
            return new OracleProperties(
                workbook.SummaryInformation?.Title,
                workbook.SummaryInformation?.Subject,
                workbook.SummaryInformation?.Author,
                workbook.SummaryInformation?.Keywords,
                workbook.SummaryInformation?.LastAuthor,
                workbook.SummaryInformation?.ApplicationName,
                workbook.DocumentSummaryInformation?.Company,
                workbook.DocumentSummaryInformation?.Manager,
                workbook.DocumentSummaryInformation?.Category);
        }
        catch (Exception)
        {
            return null;
        }
    }
}
