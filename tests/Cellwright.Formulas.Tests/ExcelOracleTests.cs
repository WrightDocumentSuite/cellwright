using System.Diagnostics;
using Cellwright.Tests.Shared;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The strongest correctness signal available offline: real workbooks carry the values Excel
/// itself computed next to each formula. Recalculating them and comparing against those
/// cached values checks the parser, the evaluator and the function catalog at once, on
/// formulas nobody on this project wrote.
/// </summary>
public sealed class ExcelOracleTests
{
    /// <summary>
    /// Share of comparable formulas that must match. The measured figure sits a little above
    /// this; what is left is dominated by SUBTOTAL over rows whose filter is no longer in the
    /// file, and by the dynamic-array and distribution functions listed in docs/formulas.md.
    /// Raise the bar as coverage grows.
    /// </summary>
    private const double RequiredAgreement = 0.96;

    private const int MaxFiles = 400;

    [Fact]
    public async Task RecalculationAgreesWithTheValuesExcelCached()
    {
        IReadOnlyList<string> files = OracleCorpus.Files(MaxFiles);
        Assert.SkipWhen(files.Count == 0, "No sample workbooks were found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        var report = new OracleReport();
        foreach (string file in files)
            await CompareFileAsync(file, report, TestContext.Current.CancellationToken);

        TestContext.Current.TestOutputHelper?.WriteLine(report.ToString());
        foreach (string line in report.TopMismatches(10).Concat(report.Samples))
            TestContext.Current.TestOutputHelper?.WriteLine(line);

        Assert.True(report.Compared > 100, $"Only {report.Compared} formulas were comparable; the corpus looks wrong.");
        Assert.True(
            report.Agreement >= RequiredAgreement,
            $"Agreement with Excel dropped to {report.Agreement:P1} (was required to stay at or above {RequiredAgreement:P0}).\n{report}");

        OracleBaseline.Enforce(report.ToBaselineLines());
    }

    private static async Task CompareFileAsync(string path, OracleReport report, CancellationToken cancellationToken)
    {
        Workbook book;
        var expected = new Dictionary<(int Sheet, int Row, int Column), CellValue>();
        try
        {
            book = await Workbook.LoadAsync(path, new LoadOptions(), cancellationToken);
            CaptureCachedValues(book, expected);
            book.Calculate(new CalculateOptions { Force = true, CancellationToken = cancellationToken });
        }
        catch (Exception exception) when (OracleCorpus.IsUnreadable(exception))
        {
            // The corpus deliberately contains damaged and legacy files from other test suites.
            report.SkippedFiles++;
            return;
        }

        report.Files++;
        for (int sheetIndex = 0; sheetIndex < book.Sheets.Count; sheetIndex++)
        {
            Worksheet sheet = book.Sheets[sheetIndex];
            foreach ((int row, int column, string? formula) in sheet.Formulas.EnumerateOccupied())
            {
                if (formula is null || !expected.TryGetValue((sheetIndex, row, column), out CellValue cached))
                    continue;

                report.Compare(cached, sheet[row, column], formula, Location(path, sheet, row, column));
            }
        }
    }

    /// <summary>Where a formula is, in a form that does not depend on where the corpus was unpacked.</summary>
    private static string Location(string path, Worksheet sheet, int row, int column) =>
        $"{OracleCorpus.RelativeName(path)}|{sheet.Name}|{new CellRef(row, column)}";

    /// <summary>Snapshots the values that were in the file before the engine overwrote them.</summary>
    private static void CaptureCachedValues(Workbook book, Dictionary<(int, int, int), CellValue> destination)
    {
        for (int sheetIndex = 0; sheetIndex < book.Sheets.Count; sheetIndex++)
        {
            Worksheet sheet = book.Sheets[sheetIndex];
            foreach ((int row, int column, string? formula) in sheet.Formulas.EnumerateOccupied())
            {
                if (formula is not null)
                    destination[(sheetIndex, row, column)] = sheet[row, column];
            }
        }
    }

    /// <summary>Counts agreements and keeps the most frequent mismatching functions.</summary>
    private sealed class OracleReport
    {
        private readonly Dictionary<string, int> _mismatchesByFunction = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _unsupportedByFunction = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _samples = [];
        private readonly HashSet<string> _sampled = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _mismatchLocations = [];
        private readonly List<string> _unsupportedLocations = [];

        /// <summary>Formulas whose cached value cannot be compared: volatile ones and cached blanks.</summary>
        public int Uncomparable { get; private set; }

        public int Files { get; set; }

        public int SkippedFiles { get; set; }

        public int Compared { get; private set; }

        public int Matched { get; private set; }

        public int Unsupported { get; private set; }

        public int OutOfScope { get; private set; }

        public double Agreement => Compared == 0 ? 1 : (double)Matched / Compared;

        /// <summary>External workbook links and the data-model functions listed in docs/formulas.md.</summary>
        private static bool IsOutOfScope(string formula) =>
            HasExternalWorkbookReference(formula) ||
            formula.Contains("CUBE", StringComparison.OrdinalIgnoreCase) ||
            formula.Contains("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) ||
            formula.Contains("FORMULATEXT", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Whether the formula points at another workbook, which xlsx writes as a bracketed
        /// index: <c>[1]Sheet1!A1</c>. Structured references also use brackets, but never around
        /// a bare number.
        /// </summary>
        private static bool HasExternalWorkbookReference(string formula)
        {
            for (int open = formula.IndexOf('['); open >= 0; open = formula.IndexOf('[', open + 1))
            {
                int close = formula.IndexOf(']', open + 1);
                if (close < 0)
                    return false;

                ReadOnlySpan<char> body = formula.AsSpan(open + 1, close - open - 1);
                if (!body.IsEmpty && !body.ContainsAnyExcept("0123456789"))
                    return true;
            }

            return false;
        }

        public void Compare(in CellValue expected, in CellValue actual, string formula, string location)
        {
            // Features documented as out of scope for v1 would otherwise drown out the signal
            // this report exists for: how well the supported surface matches Excel.
            if (IsOutOfScope(formula))
            {
                OutOfScope++;
                return;
            }

            // A #NAME? on our side means a function we have not implemented yet: that is a
            // coverage gap, tracked separately from a wrong answer.
            if (actual.IsError && actual.GetError() == ExcelError.Name && !expected.IsError)
            {
                Unsupported++;
                _unsupportedLocations.Add($"{location}|{Record(_unsupportedByFunction, formula)}");
                return;
            }

            // Volatile results and cached blanks cannot be compared meaningfully.
            if (expected.IsBlank || formula.Contains("NOW(", StringComparison.OrdinalIgnoreCase) ||
                formula.Contains("TODAY(", StringComparison.OrdinalIgnoreCase) ||
                formula.Contains("RAND", StringComparison.OrdinalIgnoreCase))
            {
                Uncomparable++;
                return;
            }

            Compared++;
            if (ValuesAgree(expected, actual))
            {
                Matched++;
                return;
            }

            string name = Record(_mismatchesByFunction, formula);
            _mismatchLocations.Add($"{location}|{name}");

            // One example per offending function rather than the first fifteen overall: a single
            // noisy function would otherwise hide every other kind of disagreement.
            if (_samples.Count < 20 && _sampled.Add(name))
                _samples.Add($"  '{formula}' expected {expected} but got {actual}");
        }

        /// <summary>A few concrete disagreements, which is what makes the report actionable.</summary>
        public IEnumerable<string> Samples => _samples;

        private static bool ValuesAgree(in CellValue expected, in CellValue actual)
        {
            if (expected.IsError || actual.IsError)
                return expected.IsError && actual.IsError && expected.GetError() == actual.GetError();

            if (expected.IsNumber && actual.IsNumber)
                return Close(expected.GetNumber(), actual.GetNumber());

            if (expected.IsDateTime || actual.IsDateTime || expected.IsTimeSpan || actual.IsTimeSpan)
                return Close(Serial(expected), Serial(actual));

            if (expected.IsBoolean || actual.IsBoolean)
                return expected.Kind == actual.Kind && expected.ToString() == actual.ToString();

            return string.Equals(expected.ToString(), actual.ToString(), StringComparison.Ordinal);
        }

        private static double Serial(in CellValue value)
        {
            try
            {
                return value.Kind switch
                {
                    CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan => value.ToSerialNumber(),
                    _ => double.NaN,
                };
            }
            catch (ArgumentOutOfRangeException)
            {
                return double.NaN;
            }
        }

        private static bool Close(double expected, double actual) =>
            Math.Abs(expected - actual) <= 1e-9 * Math.Max(1, Math.Max(Math.Abs(expected), Math.Abs(actual)));

        private static string Record(Dictionary<string, int> counters, string formula)
        {
            string name = FirstFunctionName(formula);
            counters[name] = counters.GetValueOrDefault(name) + 1;
            return name;
        }

        private static string FirstFunctionName(string formula)
        {
            int paren = formula.IndexOf('(');
            if (paren <= 0)
                return formula.Contains('[') ? "(structured references)" : "(operators)";

            int start = paren;
            while (start > 0 && (char.IsAsciiLetterOrDigit(formula[start - 1]) || formula[start - 1] is '.' or '_'))
                start--;
            if (start != paren)
                return formula[start..paren].ToUpperInvariant();

            return formula.Contains('[') ? "(structured references)" : "(operators)";
        }

        /// <summary>
        /// The result as lines that diff well. Every figure a regression could hide behind is
        /// here: the agreement percentage alone improves if comparable formulas simply vanish,
        /// and a count per function would let one new failure cancel out one fix.
        /// </summary>
        public IEnumerable<string> ToBaselineLines()
        {
            yield return $"files {Files}";
            yield return $"skipped-files {SkippedFiles}";
            yield return $"denominator {Compared + Unsupported + OutOfScope + Uncomparable}";
            yield return $"compared {Compared}";
            yield return $"matched {Matched}";
            yield return $"unsupported {Unsupported}";
            yield return $"out-of-scope {OutOfScope}";
            yield return $"uncomparable {Uncomparable}";

            foreach (string location in _mismatchLocations.Order(StringComparer.Ordinal))
                yield return $"mismatch {location}";

            foreach (string location in _unsupportedLocations.Order(StringComparer.Ordinal))
                yield return $"unsupported-function {location}";
        }

        public IEnumerable<string> TopMismatches(int count) =>
            Top(_mismatchesByFunction, count, "wrong value")
                .Concat(Top(_unsupportedByFunction, count, "unsupported"));

        private static IEnumerable<string> Top(Dictionary<string, int> counters, int count, string label) =>
            counters
                .OrderByDescending(static pair => pair.Value)
                .Take(count)
                .Select(pair => $"  {pair.Key}: {pair.Value} {label}");

        public override string ToString() =>
            $"Oracle: {Files} file(s) read, {SkippedFiles} skipped, {Compared} formula(s) compared, " +
            $"{Matched} matched ({Agreement:P1}), {Unsupported} unsupported function result(s), " +
            $"{OutOfScope} out-of-scope formula(s) ignored.";
    }
}

/// <summary>
/// Locates the sample workbooks that ship with the neighbouring spreadsheet libraries.
/// </summary>
internal static class OracleCorpus
{
    private static readonly Lazy<string[]> AllFiles = new(Discover);

    /// <summary>
    /// Whether a file simply cannot be read. Other suites keep corrupt, encrypted and
    /// hand-truncated workbooks around on purpose, and those are not our correctness signal.
    /// </summary>
    public static bool IsUnreadable(Exception exception) => exception is
        XlsxFormatException or InvalidDataException or IOException or
        NotSupportedException or InvalidOperationException or
        System.Xml.XmlException or ArgumentException;

    /// <summary>Returns up to <paramref name="max"/> sample workbooks, or an empty list.</summary>
    public static IReadOnlyList<string> Files(int max) => AllFiles.Value.Take(max).ToArray();

    /// <summary>A file's path below the corpus root, with forward slashes.</summary>
    public static string RelativeName(string path) =>
        Path.GetRelativePath(FindCorpusRoot()!, path).Replace('\\', '/');

    private static string[] Discover()
    {
        string? root = FindCorpusRoot();
        if (root is null)
            return [];

        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
        return [.. Directory.EnumerateFiles(root, "*.xlsx", options)
            .Where(CorpusRoot.Includes)
            .OrderBy(static path => path, StringComparer.Ordinal)];
    }

    /// <summary>Walks up from the test binaries looking for the folder holding the sibling repositories.</summary>
    private static string? FindCorpusRoot()
    {
        return CorpusRoot.Find("ClosedXML-0.105.0");
    }
}

/// <summary>
/// Holds the oracle to exact, reviewable results on the pinned corpus. Any change, better or
/// worse, has to be accepted by rewriting the baseline, so a new failure cannot hide behind a
/// fix and a shrinking denominator cannot pass unnoticed.
/// </summary>
internal static class OracleBaseline
{
    private const string UpdateVariable = "CELLWRIGHT_UPDATE_ORACLE_BASELINE";
    private const int MaxReportedDifferences = 40;

    public static void Enforce(IEnumerable<string> results)
    {
        // Sibling-repository discovery finds whatever happens to be unpacked, so only the
        // explicitly configured corpus has known contents to hold the result against.
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("CELLWRIGHT_CORPUS_ROOT")))
            return;

        string path = BaselinePath();
        string[] actual = [.. results];
        if (Environment.GetEnvironmentVariable(UpdateVariable) == "1")
        {
            File.WriteAllLines(path, actual);
            return;
        }

        string[] expected = File.ReadAllLines(path);
        string[] removed = [.. expected.Except(actual, StringComparer.Ordinal)];
        string[] added = [.. actual.Except(expected, StringComparer.Ordinal)];
        if (removed.Length == 0 && added.Length == 0)
            return;

        IEnumerable<string> differences = added.Select(static line => "+ " + line)
            .Concat(removed.Select(static line => "- " + line))
            .Take(MaxReportedDifferences);
        Assert.Fail(
            $"The oracle result differs from {Path.GetFileName(path)} ({added.Length} new line(s), {removed.Length} gone).\n" +
            string.Join("\n", differences) +
            $"\nIf the change is intended, review it and rerun with {UpdateVariable}=1 to accept it.");
    }

    private static string BaselinePath()
    {
        for (DirectoryInfo? directory = new(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Cellwright.slnx")))
                return Path.Combine(directory.FullName, "tests", "Cellwright.Formulas.Tests", "oracle-baseline.txt");
        }

        throw new DirectoryNotFoundException("The tests are not running from inside the repository.");
    }
}
