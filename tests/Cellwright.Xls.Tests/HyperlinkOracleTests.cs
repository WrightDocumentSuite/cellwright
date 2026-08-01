using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Checks the hyperlinks read out of the legacy format against NPOI, which unpicks the same OLE
/// monikers independently.
/// </summary>
/// <remarks>
/// A hyperlink is the one structure in the format whose fields are all optional and follow one
/// another with no marker between them, so a misread flag does not lose a field — it shifts
/// every field after it. Agreement with another implementation over the whole corpus is what
/// catches that, because a shifted read still produces a plausible-looking string.
/// </remarks>
public sealed class HyperlinkOracleTests
{
    static HyperlinkOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task HyperlinkTargetsAgreeWithTheOracle()
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

            Compare(workbook, oracle, path, ref compared, differences);
        }

        Assert.True(compared > 100, $"Only {compared} hyperlinks were comparable; coverage has dropped.");
        Assert.True(
            differences.Count == 0,
            $"{differences.Count} of {compared} hyperlinks differ from the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, differences.Take(25)));
    }

    private static void Compare(
        Workbook workbook, Dictionary<string, List<OracleLink>> oracle, string path,
        ref int compared, List<string> differences)
    {
        foreach (Worksheet sheet in workbook.Sheets)
        {
            if (!oracle.TryGetValue(sheet.Name, out List<OracleLink>? expected))
                continue;

            foreach (OracleLink link in expected)
            {
                Hyperlink? ours = sheet.Hyperlinks.Find(link.Cell);
                if (ours is null)
                    continue;

                compared++;
                if (Agrees(link, ours))
                    continue;

                differences.Add(
                    $"{Path.GetFileName(path)} [{sheet.Name}] {link.Cell}: " +
                    $"oracle '{link.Address}' vs '{ours.Target ?? ours.Location}'");
            }
        }
    }

    /// <summary>
    /// Whether the two readers name the same destination. The oracle joins the target and the
    /// place within it with a hash, and normalises separators, so the comparison is on the
    /// pieces rather than on one string.
    /// </summary>
    private static bool Agrees(OracleLink expected, Hyperlink ours)
    {
        string address = expected.Address ?? "";
        if (address.Length == 0)
            return true;

        string target = Normalize(ours.Target);
        string location = Normalize(ours.Location);
        string wanted = Normalize(address);

        if (wanted == target || wanted == location)
            return true;
        if (target.Length > 0 && location.Length > 0 && wanted == $"{target}#{location}")
            return true;

        // A relative file target is one the two spell differently often enough that only the
        // last segment is worth comparing; the leading dots are an artefact of the moniker.
        return target.Length > 0 && (wanted.EndsWith(target, StringComparison.Ordinal) ||
            target.EndsWith(wanted, StringComparison.Ordinal));
    }

    private static string Normalize(string? value) =>
        value is null ? "" : value.Replace('\\', '/').Trim();

    private readonly record struct OracleLink(CellRef Cell, string? Address);

    /// <summary>Reads the hyperlinks with NPOI, or returns null when it cannot read the file.</summary>
    private static Dictionary<string, List<OracleLink>>? ReadWithOracle(string path)
    {
        try
        {
            using FileStream stream = File.OpenRead(path);
            var workbook = new HSSFWorkbook(stream);

            var sheets = new Dictionary<string, List<OracleLink>>(StringComparer.Ordinal);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                var links = new List<OracleLink>();
                foreach (IRow row in sheet.Cast<IRow>())
                {
                    foreach (ICell cell in row.Cast<ICell>())
                    {
                        if (cell.Hyperlink is { } link)
                            links.Add(new OracleLink(new CellRef(cell.RowIndex + 1, cell.ColumnIndex + 1), link.Address));
                    }
                }

                if (links.Count > 0)
                    sheets[sheet.SheetName] = links;
            }

            return sheets;
        }
        catch (Exception)
        {
            return null;
        }
    }
}
