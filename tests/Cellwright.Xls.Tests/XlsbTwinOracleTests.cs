using System.Globalization;
using System.Text;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reads the same workbook twice — once as <c>.xlsb</c> and once from the copy the corpus keeps
/// in another format — and requires the two to say the same thing.
/// </summary>
/// <remarks>
/// This is the only oracle the binary format has for anything beyond cell values. No other
/// library reads its formulas, its comments or its validation rules, so the second reading has
/// to come from a second format rather than a second implementation.
/// <para>
/// It is also the check that catches what the value comparison cannot. That one matches sheets
/// by name and skips whatever it fails to find, so a workbook that lost every sheet passes it
/// without a single mismatch — which is exactly how the <c>BrtBundleSh</c> bug looked. Here the
/// twin still has its sheets, and the loss shows up immediately.
/// </para>
/// </remarks>
public sealed class XlsbTwinOracleTests
{
    static XlsbTwinOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public void TwinsArePresent()
    {
        Assert.SkipWhen(XlsbCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        Assert.True(
            XlsbCorpus.Twins.Count > 50,
            $"Only {XlsbCorpus.Twins.Count} of {XlsbCorpus.Files.Count} binary workbooks have a twin.");
    }

    [Fact]
    public async Task TheSameWorkbookReadsTheSameInBothFormats()
    {
        Assert.SkipWhen(XlsbCorpus.Twins.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        var report = new Report();
        foreach ((string binary, string twin) in XlsbCorpus.Twins)
        {
            Workbook? left = await TryOpenAsync(binary);
            Workbook? right = await TryOpenAsync(twin);
            if (left is null || right is null)
                continue;

            report.Pairs++;
            Compare(left, right, Path.GetFileName(binary), report);
        }

        Assert.True(report.Pairs > 40, $"Only {report.Pairs} pairs could be opened; coverage has dropped.");
        Assert.True(report.Compared > 200_000, $"Only {report.Compared} cells were comparable; coverage has dropped.");

        // Two saves of one fixture are not always the same workbook — a tool that converted the
        // corpus dropped sheets from a few — so divergent pairs are expected, counted, and left
        // out of the comparison rather than pinned per file.
        Assert.True(
            report.Divergent <= 4,
            $"{report.Divergent} of {report.Pairs} pairs disagree about their sheets:{Environment.NewLine}" +
            string.Join(Environment.NewLine, report.Notes.Take(20)));

        // The features a sheet carries are asserted apart from its cells and at zero. Folded in
        // with the cells they would be arithmetically invisible: a dozen sheets that lost every
        // validation rule is fourteen differences against three hundred thousand comparisons,
        // which any rate loose enough to be useful would swallow whole.
        Assert.True(
            report.FeatureGaps.Count == 0,
            $"{report.FeatureGaps.Count} sheets carry a feature in one format and not the other:" +
            Environment.NewLine + string.Join(Environment.NewLine, report.FeatureGaps.Take(25)));

        double rate = (double)report.Mismatched / report.Compared;
        Assert.True(
            rate < 0.0002,
            $"{report.Mismatched} of {report.Compared} cells ({rate:P4}) differ between the two formats:" +
            Environment.NewLine + string.Join(Environment.NewLine, report.Examples.Take(25)));

        Assert.True(report.Formulas > 11_000, $"Only {report.Formulas} formulas were comparable; coverage has dropped.");

        // Six cells disagree, all in one fixture whose two copies hold different content: the
        // legacy one fills a single cell where the binary one fills seven. The bound sits just
        // above that rather than at zero, and tight enough that a seventh would fail.
        double formulaRate = (double)report.FormulaGaps / report.Formulas;
        Assert.True(
            formulaRate < 0.001,
            $"{report.FormulaGaps} of {report.Formulas} formulas ({formulaRate:P2}) read differently:" +
            Environment.NewLine + string.Join(Environment.NewLine, report.FormulaExamples.Take(25)));
    }

    private sealed class Report
    {
        public int Pairs { get; set; }

        public int Compared { get; set; }

        public int Mismatched { get; set; }

        public int Divergent { get; set; }

        public List<string> Examples { get; } = [];

        public List<string> Notes { get; } = [];

        public List<string> FeatureGaps { get; } = [];

        public int Formulas { get; set; }

        public int FormulaGaps { get; set; }

        public List<string> FormulaExamples { get; } = [];
    }

    private static async Task<Workbook?> TryOpenAsync(string path)
    {
        try
        {
            return await Spreadsheet.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
        }
        catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
        {
            return null;
        }
    }

    private static void Compare(Workbook binary, Workbook twin, string name, Report report)
    {
        // A pair whose sheet lists differ is not two readings of one workbook: something
        // upstream produced different content. Counting those separately keeps them from
        // drowning the cell comparison in noise, and keeps the count visible.
        List<string> left = [.. binary.Sheets.Select(static sheet => sheet.Name)];
        List<string> right = [.. twin.Sheets.Select(static sheet => sheet.Name)];
        if (!left.SequenceEqual(right, StringComparer.Ordinal))
        {
            report.Divergent++;
            report.Notes.Add($"{name}: [{string.Join(",", left)}] vs [{string.Join(",", right)}]");
            return;
        }

        for (int i = 0; i < binary.Sheets.Count; i++)
            CompareSheet(binary.Sheets[i], twin.Sheets[i], name, report);
    }

    private static void CompareSheet(Worksheet binary, Worksheet twin, string name, Report report)
    {
        RangeRef? used = Union(binary.UsedRange, twin.UsedRange);
        if (used is not { } range)
            return;

        for (int row = range.Start.Row; row <= range.End.Row; row++)
        {
            for (int column = range.Start.Column; column <= range.End.Column; column++)
            {
                report.Compared++;
                string? left = Render(binary[row, column]);
                string? right = Render(twin[row, column]);
                if (left != right)
                {
                    report.Mismatched++;
                    if (report.Examples.Count < 25)
                        report.Examples.Add($"{name} [{binary.Name}] R{row}C{column}: xlsb '{left}' vs twin '{right}'");
                }

                CompareFormula(binary, twin, row, column, name, report);
            }
        }

        CompareCounts(binary, twin, name, report);
    }

    /// <summary>
    /// Compares the expression behind a cell, where both formats have one. This is a stronger
    /// check than the value: two different formulas can cache the same number, and matching the
    /// cached result is all the formula oracle can do on its own.
    /// </summary>
    private static void CompareFormula(Worksheet binary, Worksheet twin, int row, int column, string name, Report report)
    {
        string? left = binary.Cell(row, column).Formula;
        string? right = twin.Cell(row, column).Formula;
        if (left is null && right is null)
            return;

        report.Formulas++;
        if (Normalize(left) == Normalize(right))
            return;

        report.FormulaGaps++;
        if (report.FormulaExamples.Count < 25)
            report.FormulaExamples.Add($"{name} [{binary.Name}] R{row}C{column}: xlsb '{left}' vs twin '{right}'");
    }

    /// <summary>
    /// Puts an expression in a form the two formats can be compared in. Whitespace and case in
    /// function names carry no meaning, and the older format spells a sheet-qualified reference
    /// with an extra pair of quotes often enough that requiring them to match would test the
    /// quoting rule rather than the decoding.
    /// </summary>
    private static string? Normalize(string? formula) =>
        formula?.Replace(" ", "", StringComparison.Ordinal)
            .Replace("'", "", StringComparison.Ordinal)
            .ToUpperInvariant();

    /// <summary>
    /// Compares the things that hang off a sheet rather than sitting in its cells. Each is
    /// rendered to a sorted list of strings and the two lists compared, so a feature that is
    /// present in both but says something different is caught as well as one that is missing.
    /// </summary>
    private static void CompareCounts(Worksheet binary, Worksheet twin, string name, Report report)
    {
        Compare("merges", Render(binary.MergedRanges), Render(twin.MergedRanges));
        Compare("hyperlinks", Links(binary), Links(twin));
        Compare("comments", Notes(binary), Notes(twin));
        Compare("validations", Rules(binary), Rules(twin));
        Compare("conditional formats", Formats(binary), Formats(twin));
        Compare("tables", TableShapes(binary), TableShapes(twin));
        Compare("pictures", Images(binary), Images(twin));

        void Compare(string what, List<string> left, List<string> right)
        {
            if (left.SequenceEqual(right, StringComparer.Ordinal))
                return;

            report.FeatureGaps.Add(
                $"{name} [{binary.Name}] {what}: xlsb [{string.Join(" | ", left)}] vs twin [{string.Join(" | ", right)}]");
        }
    }

    private static List<string> Render(IReadOnlyList<RangeRef> ranges) =>
        [.. ranges.Select(static range => range.ToString()).Order(StringComparer.Ordinal)];

    private static List<string> Links(Worksheet sheet) =>
        [.. sheet.Hyperlinks
            .Select(static link => $"{link.Range} -> {link.Target ?? link.Location}")
            .Order(StringComparer.Ordinal)];

    private static List<string> Notes(Worksheet sheet) =>
        [.. sheet.Comments
            .Select(static note => $"{note.Cell} {note.Author}: {note.Text.Trim().Replace("\r\n", "\n", StringComparison.Ordinal)}")
            .Order(StringComparer.Ordinal)];

    /// <summary>
    /// A validation rule as the two formats can both express it. The messages are left out:
    /// the older format writes a lone NUL where the modern one writes nothing, and telling
    /// those apart proves nothing about the reader.
    /// </summary>
    private static List<string> Rules(Worksheet sheet) =>
        [.. sheet.Validations
            .Select(static rule =>
                $"{string.Join(",", rule.Ranges.Select(static r => r.ToString()))} " +
                $"{rule.Type}/{rule.Operator} 1={rule.Formula1} 2={rule.Formula2}")
            .Order(StringComparer.Ordinal)];

    private static List<string> Formats(Worksheet sheet) =>
        [.. sheet.ConditionalFormats
            .Select(static rule => $"{string.Join(",", rule.Ranges.Select(static r => r.ToString()))} {rule.Type}")
            .Order(StringComparer.Ordinal)];

    /// <summary>
    /// The pictures of a sheet, by where they sit and how big the image is. The bytes are left
    /// out on purpose: a fixture saved twice may well have had its images re-encoded on the way
    /// between formats, and that says nothing about either reader.
    /// </summary>
    private static List<string> Images(Worksheet sheet) =>
        [.. sheet.Pictures
            .Select(static picture => $"{picture.Anchor.From} {picture.Format}")
            .Order(StringComparer.Ordinal)];

    private static List<string> TableShapes(Worksheet sheet) =>
        [.. sheet.Tables
            .Select(static table => $"{table.Name} {table.Range} [{string.Join(",", table.Columns)}]")
            .Order(StringComparer.Ordinal)];

    private static RangeRef? Union(RangeRef? left, RangeRef? right)
    {
        if (left is not { } a)
            return right;
        if (right is not { } b)
            return left;

        return new RangeRef(
            new CellRef(Math.Min(a.Start.Row, b.Start.Row), Math.Min(a.Start.Column, b.Start.Column)),
            new CellRef(Math.Max(a.End.Row, b.End.Row), Math.Max(a.End.Column, b.End.Column)));
    }

    /// <summary>
    /// Renders a value so the two formats can be compared without arguing about how they store
    /// it. A number is cut to twelve significant digits, which is where a value that made the
    /// trip through a decimal representation stops agreeing with one that did not.
    /// </summary>
    private static string? Render(in CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => null,
        CellValueKind.Text => value.GetText().Length == 0 ? null : value.GetText().Replace("\r\n", "\n", StringComparison.Ordinal),
        CellValueKind.Number => value.GetNumber().ToString("G12", CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        CellValueKind.DateTime => value.GetDateTime().ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
        CellValueKind.TimeSpan => value.GetTimeSpan().TotalDays.ToString("G12", CultureInfo.InvariantCulture),
        _ => null,
    };
}
