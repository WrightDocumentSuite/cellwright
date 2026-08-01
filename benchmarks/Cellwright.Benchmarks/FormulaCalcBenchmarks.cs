using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Formulas;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace Cellwright.Benchmarks;

/// <summary>
/// Full recalculation of a report-shaped workbook: a data block, two derived columns, a
/// lookup against a small table and grand totals. Every engine calculates the same file.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class FormulaCalcBenchmarks
{
    private byte[] _file = [];

    [Params(50_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Cellwright Benchmarks");
        _file = BuildWorkbook(Rows);
    }

    [Benchmark(Baseline = true)]
    public double Cellwright()
    {
        using var stream = new MemoryStream(_file);
        Workbook book = Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }).GetAwaiter().GetResult();
        book.Calculate(new CalculateOptions { Force = true });
        return book.Sheets["Report"][Rows + 2, 3].GetNumber();
    }

    [Benchmark]
    public double ClosedXml()
    {
        using var stream = new MemoryStream(_file);
        using var book = new XLWorkbook(stream);
        book.RecalculateAllFormulas();
        return book.Worksheet("Report").Cell(Rows + 2, 3).GetDouble();
    }

    [Benchmark]
    public double EpPlus()
    {
        using var stream = new MemoryStream(_file);
        using var package = new ExcelPackage(stream);
        package.Workbook.Calculate();
        return Convert.ToDouble(package.Workbook.Worksheets["Report"].Cells[Rows + 2, 3].Value);
    }

    /// <summary>Writes the shared input file: constants, formulas and no cached values.</summary>
    private static byte[] BuildWorkbook(int rows)
    {
        using var stream = new MemoryStream();
        var book = Workbook.Create();
        Worksheet lookup = book.Sheets.Add("Rates");
        for (int i = 1; i <= 100; i++)
        {
            lookup[i, 1] = i;
            lookup[i, 2] = 1 + (i / 1000.0);
        }

        Worksheet report = book.Sheets.Add("Report");
        report["A1"] = "Quantity";
        report["B1"] = "Price";
        report["C1"] = "Total";
        report["D1"] = "Band";

        var random = new Random(7);
        for (int row = 2; row <= rows + 1; row++)
        {
            report[row, 1] = random.Next(1, 100);
            report[row, 2] = Math.Round(random.NextDouble() * 100, 2);
            report.Cell(row, 3).SetFormula($"A{row}*B{row}*VLOOKUP(A{row},Rates!$A$1:$B$100,2,FALSE)");
            report.Cell(row, 4).SetFormula($"IF(C{row}>2500,\"high\",IF(C{row}>1000,\"mid\",\"low\"))");
        }

        report.Cell(rows + 2, 3).SetFormula($"SUM(C2:C{rows + 1})");
        report.Cell(rows + 2, 4).SetFormula($"COUNTIF(D2:D{rows + 1},\"high\")");

        book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }).GetAwaiter().GetResult();
        return stream.ToArray();
    }
}

/// <summary>
/// Recalculation without any file handling around it. The combined load-and-calculate figure
/// is dominated by reading the package, which hides what the calculation itself costs.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 5)]
public class FormulaRecalcBenchmarks
{
    private Workbook _separateSheet = null!;
    private Workbook _sameSheet = null!;

    [Params(50_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _separateSheet = Build(Rows, tableOnOwnSheet: true);
        _sameSheet = Build(Rows, tableOnOwnSheet: false);
        _separateSheet.Calculate();
        _sameSheet.Calculate();
    }

    [Benchmark(Baseline = true, Description = "Lookup table on its own sheet")]
    public int SeparateSheet() => _separateSheet.Calculate(new CalculateOptions { Force = true }).CalculatedCells;

    /// <summary>
    /// The table lives on the sheet being written to, so every computed result lands next to
    /// it. A coarse per-sheet invalidation would throw the cached copy away on each row.
    /// </summary>
    [Benchmark(Description = "Lookup table on the sheet being written")]
    public int SameSheet() => _sameSheet.Calculate(new CalculateOptions { Force = true }).CalculatedCells;

    private static Workbook Build(int rows, bool tableOnOwnSheet)
    {
        var book = Workbook.Create();
        Worksheet report = book.Sheets.Add("Report");
        Worksheet table = tableOnOwnSheet ? book.Sheets.Add("Rates") : report;
        string reference = tableOnOwnSheet ? "Rates!$A$1:$B$100" : "$F$1:$G$100";
        int keyColumn = tableOnOwnSheet ? 1 : 6;

        for (int row = 1; row <= 100; row++)
        {
            table[row, keyColumn] = row;
            table[row, keyColumn + 1] = 1 + (row / 1000.0);
        }

        var random = new Random(7);
        for (int row = 2; row <= rows + 1; row++)
        {
            report[row, 1] = random.Next(1, 100);
            report[row, 2] = Math.Round(random.NextDouble() * 100, 2);
            report.Cell(row, 3).SetFormula($"A{row}*B{row}*VLOOKUP(A{row},{reference},2,FALSE)");
            report.Cell(row, 4).SetFormula($"IF(C{row}>2500,\"high\",IF(C{row}>1000,\"mid\",\"low\"))");
        }

        return book;
    }
}

/// <summary>
/// Cost of an edit after the workbook is already calculated: how much work the dependency
/// index saves compared with recomputing everything.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 5)]
public class IncrementalCalcBenchmarks
{
    private Workbook _chain = null!;
    private Workbook _wide = null!;

    [GlobalSetup]
    public void Setup()
    {
        _chain = BuildChain(10_000);
        _wide = BuildIndependent(50_000);
        _chain.Calculate();
        _wide.Calculate();
    }

    [Benchmark(Description = "Edit feeding a 10k-long dependency chain")]
    public int Chain()
    {
        _chain.Sheets[0]["A1"] = Random.Shared.Next(1, 100);
        return _chain.Calculate().CalculatedCells;
    }

    [Benchmark(Description = "Edit touching one of 50k independent formulas")]
    public int Independent()
    {
        _wide.Sheets[0]["A5"] = Random.Shared.Next(1, 100);
        return _wide.Calculate().CalculatedCells;
    }

    [Benchmark(Description = "Forced recalculation of the same 50k formulas")]
    public int Forced() => _wide.Calculate(new CalculateOptions { Force = true }).CalculatedCells;

    private static Workbook BuildChain(int length)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        for (int row = 2; row <= length; row++)
            sheet.Cell(row, 1).SetFormula($"A{row - 1}+1");
        return book;
    }

    private static Workbook BuildIndependent(int count)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= count; row++)
        {
            sheet[row, 1] = row;
            sheet.Cell(row, 2).SetFormula($"A{row}*2+1");
        }

        return book;
    }
}

/// <summary>
/// Aggregate throughput over a large range: the plain loop against the vectorized path used
/// once a range is big enough to pay for the copy.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 5)]
public class AggregateBenchmarks
{
    private Workbook _book = null!;

    [Params(1_000_000)]
    public int Cells { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _book = Workbook.Create();
        Worksheet sheet = _book.Sheets.Add("S");
        var random = new Random(11);
        for (int row = 1; row <= Cells; row++)
            sheet[row, 1] = random.NextDouble() * 1000;

        sheet.Cell(1, 3).SetFormula($"SUM(A1:A{Cells})");
        _book.Calculate();
    }

    [Benchmark]
    public double SumOneMillionCells() => _book.Calculate(new CalculateOptions { Force = true }).CalculatedCells;
}
