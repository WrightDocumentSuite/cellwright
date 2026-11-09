using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// Load, save, reload — over the real workbooks of the neighbouring libraries. Values must
/// survive, and so must every part of the package the model does not understand: charts,
/// pivot tables, images, printer settings, conditional formatting, macros.
/// </summary>
public sealed class PackageFidelityTests
{
    private const int MaxFiles = 400;

    /// <summary>Parts the model regenerates or deliberately drops, so their absence is not a loss.</summary>
    private static bool IsRegenerated(string entry) =>
        entry.EndsWith('/') || // Directory markers; the writer emits files only.
        entry.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
        entry.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
        entry.Contains("worksheets/", StringComparison.OrdinalIgnoreCase) ||
        entry.EndsWith("workbook.xml", StringComparison.OrdinalIgnoreCase) ||
        entry.EndsWith("styles.xml", StringComparison.OrdinalIgnoreCase) ||
        entry.EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase) ||
        entry.EndsWith("calcChain.xml", StringComparison.OrdinalIgnoreCase);

    [Fact]
    public async Task RoundTripKeepsThePartsTheModelDoesNotUnderstand()
    {
        IReadOnlyList<string> files = SampleCorpus.Files(MaxFiles);
        Assert.SkipWhen(files.Count == 0, "No sample workbooks were found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        var report = new FidelityReport();
        foreach (string file in files)
            await RoundTripAsync(file, report, TestContext.Current.CancellationToken);

        Assert.SkipWhen(report.Files == 0, "None of the sample workbooks could be read. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");
        Assert.True(report.IsClean, report.Describe());
    }

    [Fact]
    public async Task EditingACellKeepsTheChartsAndPivotTablesAroundIt()
    {
        string? path = SampleCorpus.Files(int.MaxValue)
            .FirstOrDefault(static file => file.EndsWith("ChartsheetAndPivotTable.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "The sample workbook with a chart sheet was not found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        Workbook book = await Workbook.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken);
        book.Sheets[0]["ZZ100"] = "edited";

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        HashSet<string> entries = ReadEntryNames(saved);
        Assert.Contains(entries, static entry => entry.Contains("chartsheets/", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(entries, static entry => entry.Contains("pivotCache", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(entries, static entry => entry.Contains("theme", StringComparison.OrdinalIgnoreCase));

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("edited", reloaded.Sheets[0]["ZZ100"].GetText());
    }

    [Fact]
    public async Task PreservationCanBeTurnedOff()
    {
        string? path = SampleCorpus.Files(int.MaxValue)
            .FirstOrDefault(static file => file.EndsWith("ChartsheetAndPivotTable.xlsx", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "The sample workbook with a chart sheet was not found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        var options = new LoadOptions { PreserveUnknownParts = false };
        Workbook book = await Workbook.LoadAsync(path!, options, TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        HashSet<string> entries = ReadEntryNames(saved);
        Assert.DoesNotContain(entries, static entry => entry.Contains("pivotCache", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(entries, static entry => entry.Contains("workbook.xml", StringComparison.OrdinalIgnoreCase));
    }

    private static async Task RoundTripAsync(string path, FidelityReport report, CancellationToken cancellationToken)
    {
        Workbook book;
        HashSet<string> entries;
        try
        {
            book = await Workbook.LoadAsync(path, cancellationToken: cancellationToken);
            entries = ReadEntryNames(path);
        }
        catch (Exception exception) when (SampleCorpus.IsUnreadable(exception))
        {
            // The corpus deliberately contains damaged and legacy files from other test suites.
            report.Skipped++;
            return;
        }

        report.Files++;
        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, cancellationToken);

        // A Strict package reads but does not round-trip: the writer is transitional, and the
        // loader says so rather than producing a file that is half of each conformance class.
        if (book.LoadDiagnostics?.Warnings.Any(static w => w.Code == "StrictConformance") == true)
        {
            report.StrictFiles++;
            return;
        }

        string name = Path.GetFileName(path);
        saved.Position = 0;
        foreach (string missing in entries.Except(ReadEntryNames(saved), StringComparer.OrdinalIgnoreCase))
        {
            if (!IsRegenerated(missing))
                report.LostParts.Add($"{name}: {missing}");
        }

        // Only errors we introduce count: the corpus is full of files that were already
        // invalid when their producer wrote them, and preserving a defect faithfully is
        // exactly the behaviour under test.
        saved.Position = 0;
        HashSet<string> known = ValidateFile(path, cancellationToken);
        foreach (string introduced in Validate(saved, cancellationToken).Except(known, StringComparer.Ordinal))
            report.Invalid.Add($"{name}: {introduced}");

        using (Stream source = File.OpenRead(path))
        {
            HashSet<string> knownProblems = PackageInvariants.Problems(source);
            foreach (string introduced in PackageInvariants.Problems(saved).Except(knownProblems, StringComparer.Ordinal))
                report.Invalid.Add($"{name}: {introduced}");
        }

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(saved, new LoadOptions { LeaveOpen = true }, cancellationToken);
        CompareValues(name, book, reloaded, report);
    }

    private static void CompareValues(string name, Workbook original, Workbook reloaded, FidelityReport report)
    {
        if (original.Sheets.Count != reloaded.Sheets.Count)
        {
            report.ValueMismatches.Add($"{name}: {original.Sheets.Count} sheet(s) became {reloaded.Sheets.Count}");
            return;
        }

        for (int i = 0; i < original.Sheets.Count && report.ValueMismatches.Count < 20; i++)
        {
            Worksheet source = original.Sheets[i];
            Worksheet target = reloaded.Sheets[i];
            if (source.UsedRange is not { } used)
                continue;

            for (int row = used.Start.Row; row <= used.End.Row; row++)
            {
                for (int column = used.Start.Column; column <= used.End.Column; column++)
                {
                    CellValue expected = source[row, column];
                    CellValue actual = target[row, column];
                    report.Cells++;
                    if (expected != actual)
                        report.ValueMismatches.Add($"{name}!{source.Name}!{new CellRef(row, column)}: {expected} became {actual}");
                }
            }
        }
    }

    /// <summary>Validation errors of a package, as stable strings.</summary>
    private static HashSet<string> Validate(Stream stream, CancellationToken cancellationToken)
    {
        try
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
            return
            [
                .. new OpenXmlValidator().Validate(document, cancellationToken)
                    .Select(static error => $"{error.Path?.XPath} {error.Description}")
                    .Take(200),
            ];
        }
        catch (Exception exception) when (SampleCorpus.IsUnreadable(exception))
        {
            return [$"{exception.GetType().Name} {exception.Message}"];
        }
    }

    private static HashSet<string> ValidateFile(string path, CancellationToken cancellationToken)
    {
        using FileStream stream = File.OpenRead(path);
        return Validate(stream, cancellationToken);
    }

    private static HashSet<string> ReadEntryNames(string path)
    {
        using FileStream stream = File.OpenRead(path);
        return ReadEntryNames(stream);
    }

    private static HashSet<string> ReadEntryNames(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true, Encoding.UTF8);
        return [.. archive.Entries.Select(static entry => entry.FullName)];
    }

    private sealed class FidelityReport
    {
        public int Files { get; set; }

        public int Skipped { get; set; }

        /// <summary>Strict-conformance packages, which are read but deliberately not round-tripped.</summary>
        public int StrictFiles { get; set; }

        public int Cells { get; set; }

        public List<string> LostParts { get; } = [];

        public List<string> Invalid { get; } = [];

        public List<string> ValueMismatches { get; } = [];

        public bool IsClean => LostParts.Count == 0 && Invalid.Count == 0 && ValueMismatches.Count == 0;

        public string Describe() =>
            $"{this}\n{Section("Lost parts", LostParts)}{Section("Invalid packages", Invalid)}{Section("Value mismatches", ValueMismatches)}";

        public override string ToString() =>
            $"Fidelity: {Files} file(s) round-tripped, {Skipped} skipped, {StrictFiles} strict, " +
            $"{Cells} cell(s) compared, {LostParts.Count} lost part(s), {Invalid.Count} invalid package(s), " +
            $"{ValueMismatches.Count} value mismatch(es).";

        private static string Section(string title, List<string> items) =>
            items.Count == 0 ? "" : $"{title}:\n  {string.Join("\n  ", items.Take(15))}\n";
    }
}
