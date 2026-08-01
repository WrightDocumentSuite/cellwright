using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Formulas;
using Cellwright.Formulas.Engine;

namespace Cellwright.Benchmarks;

/// <summary>
/// The subscription lookup on its own, against the linear scan it replaced. Whole-workbook
/// benchmarks bury this behind the cost of recalculating the formulas that were found.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Throughput, launchCount: 1, warmupCount: 2, iterationCount: 5)]
public class SubscriptionLookupBenchmarks
{
    private readonly HashSet<int> _found = [];
    private RangeSubscriptionIndex _index = null!;
    private (int FirstRow, int LastRow, int FirstColumn, int LastColumn, int Entry)[] _linear = [];

    [Params(50_000)]
    public int Subscriptions { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var builder = new RangeSubscriptionIndex.Builder();
        _linear = new (int, int, int, int, int)[Subscriptions];
        for (int entry = 0; entry < Subscriptions; entry++)
        {
            // Running-total shape: every range starts at row one and ends a row lower.
            int lastRow = entry + 1;
            builder.Add(new RangeKey(0, 1, 1, lastRow, 1), entry);
            _linear[entry] = (1, lastRow, 1, 1, entry);
        }

        _index = builder.Build();
    }

    [Benchmark(Baseline = true, Description = "Interval index")]
    public int Indexed()
    {
        _found.Clear();
        _index.Collect(0, Subscriptions, 1, _found);
        return _found.Count;
    }

    [Benchmark(Description = "Linear scan over the same subscriptions")]
    public int Linear()
    {
        _found.Clear();
        foreach ((int firstRow, int lastRow, int firstColumn, int lastColumn, int entry) in _linear)
        {
            if (Subscriptions >= firstRow && Subscriptions <= lastRow && firstColumn <= 1 && lastColumn >= 1)
                _found.Add(entry);
        }

        return _found.Count;
    }
}

/// <summary>
/// Cost of an edit in a workbook whose formulas read ranges rather than single cells. The
/// existing incremental benchmark only uses point references, so it never exercised the
/// structure that finds range subscriptions.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 5)]
public class RangeEditBenchmarks
{
    private Workbook _sharedTable = null!;
    private Workbook _runningTotals = null!;

    [Params(50_000)]
    public int Formulas { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _sharedTable = BuildSharedTable(Formulas);
        _runningTotals = BuildRunningTotals(Formulas);
        _sharedTable.Calculate();
        _runningTotals.Calculate();
    }

    [Benchmark(Description = "Edit outside every subscribed range")]
    public int EditOutsideRanges()
    {
        Worksheet sheet = _sharedTable.Sheets["Report"];
        sheet["H1"] = Random.Shared.Next(1, 100);
        return _sharedTable.Calculate().CalculatedCells;
    }

    [Benchmark(Description = "Edit inside a table read by 50k lookups")]
    public int EditInsideSharedTable()
    {
        Worksheet rates = _sharedTable.Sheets["Rates"];
        rates[50, 2] = 1 + (Random.Shared.Next(1, 100) / 1000.0);
        return _sharedTable.Calculate().CalculatedCells;
    }

    /// <summary>
    /// Every running total covers a different prefix, so the index has to sort through fifty
    /// thousand distinct ranges. Touching the last row keeps the recalculation itself down to
    /// a single formula, leaving the lookup as what is being measured.
    /// </summary>
    [Benchmark(Description = "Edit found among 50k distinct ranges")]
    public int EditAmongDistinctRanges()
    {
        Worksheet sheet = _runningTotals.Sheets["S"];
        sheet[Formulas, 1] = Random.Shared.Next(1, 100);
        return _runningTotals.Calculate().CalculatedCells;
    }

    /// <summary>Every formula reads the same lookup table: the subscriptions all coincide.</summary>
    private static Workbook BuildSharedTable(int rows)
    {
        var book = Workbook.Create();
        Worksheet rates = book.Sheets.Add("Rates");
        for (int row = 1; row <= 100; row++)
        {
            rates[row, 1] = row;
            rates[row, 2] = 1 + (row / 1000.0);
        }

        Worksheet report = book.Sheets.Add("Report");
        var random = new Random(3);
        for (int row = 1; row <= rows; row++)
        {
            report[row, 1] = random.Next(1, 100);
            report[row, 2] = random.NextDouble() * 100;
            report.Cell(row, 3).SetFormula($"A{row}*B{row}*VLOOKUP(A{row},Rates!$A$1:$B$100,2,FALSE)");
        }

        return book;
    }

    /// <summary>Each formula reads a different prefix, so no two subscriptions are the same.</summary>
    private static Workbook BuildRunningTotals(int rows)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= rows; row++)
        {
            sheet[row, 1] = row;
            sheet.Cell(row, 2).SetFormula($"SUM(A$1:A{row})");
        }

        return book;
    }
}
