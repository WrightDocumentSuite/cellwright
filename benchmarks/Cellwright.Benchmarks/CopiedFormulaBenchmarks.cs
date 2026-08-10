using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Formulas;

namespace Cellwright.Benchmarks;

/// <summary>
/// Conditional-format shaped formula evaluation: one stored formula copied down a large range.
/// The source text remains the program-cache key and reference deltas are applied at execution.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[InvocationCount(5)]
public class CopiedFormulaBenchmarks
{
    private Worksheet _sheet = null!;
    private static readonly CellRef Source = new(1, 2);

    [Params(100_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Workbook book = Workbook.Create();
        _sheet = book.Sheets.Add("Report");
        for (int row = 1; row <= Rows; row++)
            _sheet[row, 1] = row & 1;
    }

    [Benchmark]
    public int RelativeExpression()
    {
        int matched = 0;
        for (int row = 1; row <= Rows; row++)
        {
            if (_sheet.EvaluateCopied("$A1>0", Source, new CellRef(row, 2)).GetBoolean())
                matched++;
        }

        return matched;
    }
}

/// <summary>
/// Sliding 32-row lookup ranges deliberately cross the historical two-million-retained-cell
/// boundary. This catches eviction implementations whose nominal LRU operation hides a scan.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class CopiedRangeCacheBenchmarks
{
    private Worksheet _sheet = null!;
    private static readonly CellRef Source = new(1, 4);

    [Params(70_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Workbook book = Workbook.Create();
        _sheet = book.Sheets.Add("Report");
        for (int row = 1; row <= Rows + 31; row++)
        {
            _sheet[row, 1] = row;
            _sheet[row, 2] = row * 10;
        }
    }

    [Benchmark]
    public double SlidingVlookup()
    {
        double checksum = 0;
        for (int row = 1; row <= Rows; row++)
        {
            checksum += _sheet.EvaluateCopied(
                "VLOOKUP(A1,A1:B32,2,FALSE)",
                Source,
                new CellRef(row, 4)).GetNumber();
        }

        return checksum;
    }
}
