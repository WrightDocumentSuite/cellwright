using System.Text;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reads every binary workbook the sibling repositories keep and checks the values against an
/// independent reader of the same format, exactly as the legacy corpus is checked. There is no
/// second implementation of the formula text to compare against, so formulas are held to the
/// weaker but still useful standard of recomputing to the value the file cached.
/// </summary>
public sealed class XlsbCorpusTests
{
    static XlsbCorpusTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public void CorpusIsPresent()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No sibling repositories with .xlsb test data were found. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        Assert.True(XlsbCorpus.Files.Count > 40, $"Expected a sizeable corpus, found {XlsbCorpus.Files.Count} files.");
    }

    /// <summary>
    /// Every file either loads or fails with an exception that says why. What must not happen
    /// is an exception from the middle of the reader — a null reference, an index out of range
    /// — because that means a malformed file reached code that assumed it was not.
    /// </summary>
    [Fact]
    public async Task EveryWorkbookLoadsOrFailsCleanly()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var unexpected = new List<string>();
        int loaded = 0;

        foreach (string path in XlsbCorpus.Files)
        {
            try
            {
                Workbook workbook = await Spreadsheet.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
                Assert.NotNull(workbook);
                loaded++;
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                // An encrypted file with no password, or one the corpus keeps because it is broken.
            }
            catch (Exception exception)
            {
                unexpected.Add($"{Path.GetFileName(path)}: {exception.GetType().Name}: {exception.Message}");
            }
        }

        Assert.True(
            unexpected.Count == 0,
            $"{unexpected.Count} of {XlsbCorpus.Files.Count} workbooks failed unexpectedly:{Environment.NewLine}" +
            string.Join(Environment.NewLine, unexpected.Take(25)));
        Assert.True(loaded > XlsbCorpus.Files.Count / 2, $"Only {loaded} of {XlsbCorpus.Files.Count} workbooks loaded.");
    }

    /// <summary>
    /// Every sheet the oracle sees is in the workbook this reader built. The value comparison
    /// matches sheets by name and skips what it cannot find, so a workbook that lost all its
    /// sheets would sail through it — which is exactly what a mis-read sheet list looks like.
    /// </summary>
    [Fact]
    public async Task EverySheetTheOracleSeesIsRead()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var missing = new List<string>();
        foreach (string path in XlsbCorpus.Files)
        {
            Workbook workbook;
            try
            {
                workbook = await Spreadsheet.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            if (CorpusTests.ReadWithOracle(path) is not { } oracle)
                continue;

            foreach (CorpusTests.OracleSheet expected in oracle)
            {
                if (workbook.Sheets.Find(expected.Name) is null)
                    missing.Add($"{Path.GetFileName(path)}: [{expected.Name}]");
            }
        }

        Assert.True(
            missing.Count == 0,
            $"{missing.Count} sheets are missing from the workbooks that were read:{Environment.NewLine}" +
            string.Join(Environment.NewLine, missing.Take(25)));
    }

    /// <summary>Compares the cell values against ExcelDataReader over the whole corpus.</summary>
    [Fact]
    public async Task CellValuesAgreeWithTheOracle()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int compared = 0;
        int mismatched = 0;
        var examples = new List<string>();

        foreach (string path in XlsbCorpus.Files)
        {
            Workbook workbook;
            try
            {
                workbook = await Spreadsheet.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            CorpusTests.OracleSheet[]? oracle = CorpusTests.ReadWithOracle(path);
            if (oracle is null)
                continue;

            CorpusTests.CompareWorkbook(workbook, oracle, path, ref compared, ref mismatched, examples);
        }

        // The floor is not ceremony. A cell this reader never produced is not a mismatch — it
        // is a cell the comparison never reached — so a bug that loses whole sheets shows up
        // here as a drop in coverage and nowhere else.
        Assert.True(compared > 250_000, $"Only {compared} cells were comparable; coverage has dropped.");

        // The two readers agree about every cell of this corpus. Unlike the legacy format,
        // where a handful of files are genuinely ambiguous, there is nothing here to argue
        // about, so the bound is exact rather than a rate.
        Assert.True(
            mismatched == 0,
            $"{mismatched} of {compared} cells differ from the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, examples.Take(25)));
    }
}
