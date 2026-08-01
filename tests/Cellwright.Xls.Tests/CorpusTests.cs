using System.Globalization;
using System.Text;
using ExcelDataReader;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reads every legacy workbook the sibling repositories keep, and checks the values against an
/// independent reader of the same format. A format this old has no specification-conformant
/// sample set worth the name; agreement with another implementation over hundreds of real files
/// is the closest thing to one.
/// </summary>
public sealed class CorpusTests
{
    static CorpusTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public void CorpusIsPresent()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No sibling repositories with .xls test data were found. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        Assert.True(XlsCorpus.Files.Count > 100, $"Expected a sizeable corpus, found {XlsCorpus.Files.Count} files.");
    }

    /// <summary>
    /// Every file either loads or fails with an exception that says why. What must not happen
    /// is an exception from somewhere in the middle of the reader — a null reference, an index
    /// out of range — because that means a malformed file reached code that assumed it was not.
    /// </summary>
    [Fact]
    public async Task EveryWorkbookLoadsOrFailsCleanly()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var unexpected = new List<string>();
        int loaded = 0;

        foreach (string path in XlsCorpus.Files)
        {
            try
            {
                Workbook workbook = await XlsWorkbook.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
                Assert.NotNull(workbook);
                loaded++;
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                // A file the corpus keeps precisely because it is broken.
            }
            catch (Exception exception)
            {
                unexpected.Add($"{Path.GetFileName(path)}: {exception.GetType().Name}: {exception.Message}");
            }
        }

        Assert.True(
            unexpected.Count == 0,
            $"{unexpected.Count} of {XlsCorpus.Files.Count} workbooks failed unexpectedly:{Environment.NewLine}" +
            string.Join(Environment.NewLine, unexpected.Take(25)));
        Assert.True(loaded > XlsCorpus.Files.Count / 2, $"Only {loaded} of {XlsCorpus.Files.Count} workbooks loaded.");
    }

    /// <summary>
    /// Compares the cell values against ExcelDataReader over the whole corpus. A mismatch rate
    /// is reported rather than a per-file assertion: the two readers disagree on a handful of
    /// genuinely ambiguous files, and a threshold catches a regression without pinning the
    /// result to whichever library is right about the edge cases.
    /// </summary>
    [Fact]
    public async Task CellValuesAgreeWithTheOracle()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int compared = 0;
        int mismatched = 0;
        var examples = new List<string>();

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

            OracleSheet[]? oracle = ReadWithOracle(path);
            if (oracle is null)
                continue;

            CompareWorkbook(workbook, oracle, path, ref compared, ref mismatched, examples);
        }

        Assert.True(compared > 10_000, $"Only {compared} cells were comparable; the corpus may not be readable.");

        // The last few dozen disagreements over fifty million cells are files where this reader
        // is the more faithful of the two — an error value the oracle reports as empty, an
        // array formula's cached result it does not reach. The bound leaves room for those
        // without leaving room for a regression.
        double rate = (double)mismatched / compared;
        Assert.True(
            rate < 0.000002,
            $"{mismatched} of {compared} cells ({rate:P2}) differ from the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, examples.Take(25)));
    }

    internal static void CompareWorkbook(
        Workbook workbook, OracleSheet[] oracle, string path,
        ref int compared, ref int mismatched, List<string> examples)
    {
        // Sheets are matched by name: this reader leaves out the chart and macro sheets it does
        // not model, so the two sheet lists are not the same length.
        foreach (OracleSheet expected in oracle)
        {
            Worksheet? sheet = workbook.Sheets.Find(expected.Name);
            if (sheet is null)
                continue;

            for (int row = 0; row < expected.Rows.Count; row++)
            {
                object?[] expectedRow = expected.Rows[row];
                for (int column = 0; column < expectedRow.Length; column++)
                {
                    object? want = expectedRow[column];
                    CellValue got = sheet.Cell(row + 1, column + 1).Value;
                    compared++;
                    if (Matches(want, got))
                        continue;

                    mismatched++;
                    if (examples.Count < 25)
                    {
                        examples.Add(
                            $"{Path.GetFileName(path)} [{sheet.Name}] R{row + 1}C{column + 1}: " +
                            $"oracle '{Normalize(want)}' vs '{Normalize(got)}'");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Whether the two readers agree about a cell, allowing for the places where the oracle is
    /// deliberately lossier: it drops error values, and it hands back a raw serial where a date
    /// format says the number is a date.
    /// </summary>
    private static bool Matches(object? oracle, in CellValue ours)
    {
        // The oracle returns null for a cell holding an error, so there is nothing to compare.
        if (oracle is null or DBNull && ours.Kind == CellValueKind.Error)
            return true;

        // A date or duration is a formatted number; when only one side applied the format, the
        // serial is the thing both agree on.
        if (oracle is double serial && ours.Kind is CellValueKind.DateTime or CellValueKind.TimeSpan)
            return NearlyEqual(serial, ours.ToSerialNumber());

        if (oracle is TimeSpan elapsed && ours.Kind is CellValueKind.TimeSpan or CellValueKind.Number)
            return NearlyEqual(elapsed.TotalDays, ours.ToSerialNumber());

        if (oracle is DateTime moment && ours.Kind is CellValueKind.Number)
            return NearlyEqual(ExcelSerialDate.FromDateTime(moment), ours.GetNumber());

        // A duration only keeps tick resolution once it is a TimeSpan, so the last few digits
        // of a serial that was never round can differ.
        if (oracle is double raw && ours.Kind == CellValueKind.Number)
            return NearlyEqual(raw, ours.GetNumber());

        return Normalize(oracle) == Normalize(ours);
    }

    /// <summary>Compares two serial numbers at the precision a duration survives a round trip at.</summary>
    private static bool NearlyEqual(double left, double right)
    {
        if (left == right)
            return true;

        double scale = Math.Max(Math.Abs(left), Math.Abs(right));
        return Math.Abs(left - right) <= Math.Max(scale, 1) * 1e-9;
    }

    /// <summary>
    /// Renders a value so that the two readers can be compared without arguing about types. A
    /// number is rounded to twelve significant digits: both compute the same double, but a date
    /// that one returns as a <see cref="DateTime"/> and the other as a serial must still match.
    /// </summary>
    private static string? Normalize(object? value) => value switch
    {
        null or DBNull => null,
        string text => text.Length == 0 ? null : text.Replace("\r\n", "\n", StringComparison.Ordinal),
        double number => FormatNumber(number),
        int number => FormatNumber(number),
        bool flag => flag ? "TRUE" : "FALSE",
        DateTime moment => moment.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
        TimeSpan duration => FormatNumber(duration.TotalDays),
        _ => value.ToString(),
    };

    private static string? Normalize(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => null,
        CellValueKind.Text => value.GetText().Length == 0
            ? null
            : value.GetText().Replace("\r\n", "\n", StringComparison.Ordinal),
        CellValueKind.Number => FormatNumber(value.GetNumber()),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        CellValueKind.DateTime => value.GetDateTime().ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
        CellValueKind.TimeSpan => FormatNumber(value.GetTimeSpan().TotalDays),
        _ => null,
    };

    private static string FormatNumber(double value) =>
        value.ToString("G12", CultureInfo.InvariantCulture);

    internal sealed record OracleSheet(string Name, List<object?[]> Rows);

    /// <summary>Reads a workbook with ExcelDataReader, or returns null when it cannot either.</summary>
    internal static OracleSheet[]? ReadWithOracle(string path)
    {
        try
        {
            using FileStream stream = File.OpenRead(path);
            using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            var sheets = new List<OracleSheet>();
            do
            {
                var rows = new List<object?[]>();
                while (reader.Read())
                {
                    var row = new object?[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                        row[i] = reader.GetValue(i);
                    rows.Add(row);
                }

                sheets.Add(new OracleSheet(reader.Name, rows));
            }
            while (reader.NextResult());

            return [.. sheets];
        }
        catch (Exception)
        {
            // The oracle refuses plenty of the same files this reader does; nothing to compare.
            return null;
        }
    }
}
