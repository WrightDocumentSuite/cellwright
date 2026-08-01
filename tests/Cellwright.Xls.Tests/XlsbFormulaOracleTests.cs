using System.Text;
using Cellwright.Formulas;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Checks the formula decoder against the workbook's own cached results, the same way the
/// legacy corpus is checked.
/// </summary>
/// <remarks>
/// This is the only check available for the binary format's formulas. No other library reads
/// them back as text — the one used as a value oracle does not decode formulas at all — so
/// there is nothing to compare a rendered expression against. Recomputing it and matching the
/// value the file cached is weaker, since two different expressions can agree on one result,
/// but it catches a misread token stream immediately: a reference off by a row does not
/// evaluate to the same number.
/// </remarks>
public sealed class XlsbFormulaOracleTests
{
    static XlsbFormulaOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task DecodedFormulasEvaluateToTheirCachedResults()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int agreed = 0;
        int total = 0;
        var examples = new List<string>();
        var byFile = new Dictionary<string, int>(StringComparer.Ordinal);

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

            FormulaOracleTests.Compare(workbook, path, ref agreed, ref total, examples, byFile);
        }

        // A decoder that returned nothing would agree with nothing and pass a rate test by
        // having no cases, so the count is asserted before the rate.
        Assert.True(total > 11_000, $"Only {total} formulas were comparable; coverage has dropped.");

        string worst = string.Join(
            Environment.NewLine,
            byFile.OrderByDescending(static pair => pair.Value).Take(10).Select(static pair => $"  {pair.Key}: {pair.Value}"));

        // Every formula in the corpus currently recomputes to the value the file cached. The
        // bound leaves room for a workbook saved stale rather than pinning it at exactly all.
        double rate = (double)agreed / total;
        Assert.True(
            rate > 0.999,
            $"Only {agreed} of {total} formulas ({rate:P1}) evaluate to their cached result." +
            $"{Environment.NewLine}Worst files:{Environment.NewLine}{worst}{Environment.NewLine}" +
            string.Join(Environment.NewLine, examples.Take(25)));
    }
}
