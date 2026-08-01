using System.Globalization;
using System.Text;
using Cellwright.Formulas;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Checks the formula decoder against the workbook's own cached results. Every formula cell in
/// a .xls file carries the value Excel last computed for it, which makes each one a test case
/// nobody had to write: decode the expression, evaluate it, and the two should agree.
/// </summary>
/// <remarks>
/// The two can legitimately differ. A formula may depend on the date, on a sheet this library
/// skipped, or on a function it does not implement, and the file may simply have been saved
/// without recalculating. So the measure is an agreement rate over the whole corpus rather than
/// a pass or fail per formula.
/// </remarks>
public sealed class FormulaOracleTests
{
    static FormulaOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task DecodedFormulasEvaluateToTheirCachedResults()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int agreed = 0;
        int total = 0;
        var examples = new List<string>();
        var byFile = new Dictionary<string, int>(StringComparer.Ordinal);

        // A slice of the corpus: recalculating every workbook takes minutes, and the point is a
        // broad signal rather than exhaustive coverage.
        foreach (string path in XlsCorpus.Take(150))
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

            Compare(workbook, path, ref agreed, ref total, examples, byFile);
        }

        Assert.True(total > 500, $"Only {total} formulas were comparable; the corpus may not hold enough.");

        // Most of what is left is a workbook saved without recalculating, so the count per file
        // is worth seeing: a decoder bug spreads across files, a stale cache does not.
        string worst = string.Join(
            Environment.NewLine,
            byFile.OrderByDescending(static pair => pair.Value).Take(10).Select(static pair => $"  {pair.Key}: {pair.Value}"));

        double rate = (double)agreed / total;
        Assert.True(
            rate > 0.97,
            $"Only {agreed} of {total} formulas ({rate:P1}) evaluate to their cached result." +
            $"{Environment.NewLine}Worst files:{Environment.NewLine}{worst}{Environment.NewLine}" +
            string.Join(Environment.NewLine, examples.Take(25)));
    }

    internal static void Compare(
        Workbook workbook, string path, ref int agreed, ref int total,
        List<string> examples, Dictionary<string, int> byFile)
    {
        // The cached values have to be captured before recalculating, since that overwrites them.
        var cached = new List<(Worksheet Sheet, CellRef Cell, string Formula, CellValue Value)>();
        foreach (Worksheet sheet in workbook.Sheets)
        {
            foreach ((int row, int column, string? formula) in sheet.Formulas.EnumerateOccupied())
            {
                if (formula is null)
                    continue;

                if (IsUnreproducible(formula))
                    continue;

                var cell = new CellRef(row, column);
                cached.Add((sheet, cell, formula, sheet[row, column]));
            }
        }

        if (cached.Count == 0)
            return;

        try
        {
            workbook.Calculate();
        }
        catch (Exception exception) when (exception is InvalidOperationException or NotSupportedException)
        {
            return;
        }

        foreach ((Worksheet sheet, CellRef cell, string formula, CellValue expected) in cached)
        {
            total++;
            CellValue actual = sheet[cell.Row, cell.Column];
            if (Agrees(expected, actual))
            {
                agreed++;
                continue;
            }

            string name = Path.GetFileName(path);
            byFile[name] = byFile.GetValueOrDefault(name) + 1;
            if (examples.Count < 25)
            {
                examples.Add(
                    $"{name} [{sheet.Name}] {cell} ={formula}: cached '{Render(expected)}' " +
                    $"but computed '{Render(actual)}'");
            }
        }
    }

    /// <summary>
    /// Whether a formula's cached value cannot be reproduced by recalculating, whatever the
    /// decoder does. A volatile function answers differently every time it is asked, and a
    /// reference into a deleted range or another workbook has nothing left to read.
    /// </summary>
    private static bool IsUnreproducible(string formula)
    {
        if (formula.Contains("#REF!", StringComparison.Ordinal) ||
            formula.Contains('[', StringComparison.Ordinal))
        {
            return true;
        }

        foreach (string name in Volatile)
        {
            if (formula.Contains(name, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static readonly string[] Volatile =
    [
        "TODAY(", "NOW(", "RAND(", "RANDBETWEEN(", "OFFSET(", "INDIRECT(", "INFO(", "CELL(",
    ];

    /// <summary>
    /// Whether a recomputed value matches what the file cached, allowing for the last digits of
    /// a double and for the two ways an empty result can be spelled.
    /// </summary>
    private static bool Agrees(in CellValue expected, in CellValue actual)
    {
        if (expected.Kind != actual.Kind)
        {
            // A blank and an empty string mean the same thing at the end of a formula.
            bool expectedEmpty = expected.IsBlank || (expected.IsText && expected.GetText().Length == 0);
            bool actualEmpty = actual.IsBlank || (actual.IsText && actual.GetText().Length == 0);
            if (expectedEmpty && actualEmpty)
                return true;

            return IsNumeric(expected) && IsNumeric(actual) && NearlyEqual(expected, actual);
        }

        return expected.Kind switch
        {
            CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan => NearlyEqual(expected, actual),
            CellValueKind.Text => string.Equals(expected.GetText(), actual.GetText(), StringComparison.Ordinal),
            CellValueKind.Boolean => expected.GetBoolean() == actual.GetBoolean(),
            CellValueKind.Error => expected.GetError() == actual.GetError(),
            _ => true,
        };
    }

    private static bool IsNumeric(in CellValue value) =>
        value.Kind is CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan;

    private static bool NearlyEqual(in CellValue left, in CellValue right)
    {
        double a = left.ToSerialNumber();
        double b = right.ToSerialNumber();
        if (a == b)
            return true;
        if (double.IsNaN(a) || double.IsNaN(b))
            return false;

        double scale = Math.Max(Math.Abs(a), Math.Abs(b));
        return Math.Abs(a - b) <= Math.Max(scale, 1) * 1e-9;
    }

    private static string Render(in CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Text => value.GetText(),
        CellValueKind.Number => value.GetNumber().ToString("G12", CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        CellValueKind.DateTime => value.GetDateTime().ToString("O", CultureInfo.InvariantCulture),
        CellValueKind.TimeSpan => value.GetTimeSpan().ToString(),
        _ => "?",
    };
}
