using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Formulas.Parsing;

namespace Cellwright.Benchmarks;

/// <summary>
/// Parsing throughput on a mix of formulas typical of report workbooks. There is no
/// third-party baseline here on purpose: the parser ClosedXML ships expects a visitor with
/// two dozen members, so a comparison would mostly measure the adapter written for it.
/// The numbers that matter are absolute throughput and how much the program cache saves,
/// since real sheets repeat the same formula down whole columns.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Throughput, launchCount: 1, warmupCount: 2, iterationCount: 5)]
public class FormulaParseBenchmarks
{
    private string[] _distinct = [];
    private string[] _repeating = [];
    private Workbook _workbook = null!;

    [Params(10_000)]
    public int Count { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _workbook = Workbook.Create();
        _workbook.Sheets.Add("Sheet1");
        _workbook.Sheets.Add("Data");

        string[] templates =
        [
            "A{0}*B{0}", "SUM(A1:A{0})", "IF(A{0}>100,\"high\",\"low\")",
            "VLOOKUP($A{0},Data!$A$1:$D$999,3,FALSE)", "ROUND(A{0}*(1+$B$1),2)",
            "SUMIFS(Data!$C:$C,Data!$A:$A,A{0},Data!$B:$B,\">0\")",
            "(A{0}+B{0})/MAX(C{0},1)-MIN(D{0}:E{0})",
        ];

        _distinct = new string[Count];
        _repeating = new string[Count];
        for (int i = 0; i < Count; i++)
        {
            _distinct[i] = string.Format(templates[i % templates.Length], (i % 900) + 1);
            _repeating[i] = string.Format(templates[i % templates.Length], 1);
        }
    }

    [Benchmark(Baseline = true, Description = "Parse 10k distinct formulas")]
    public int Distinct()
    {
        int length = 0;
        foreach (string formula in _distinct)
            length += FormulaParser.Parse(formula, _workbook).Instructions.Length;
        return length;
    }

    [Benchmark(Description = "Parse 10k formulas of 7 distinct shapes through the cache")]
    public int Cached()
    {
        var cache = new ProgramCache(_workbook);
        int length = 0;
        foreach (string formula in _repeating)
            length += cache.GetOrParse(formula, out _)!.Instructions.Length;
        return length;
    }

    [Benchmark(Description = "Render 10k parsed programs back to A1 text")]
    public int Render()
    {
        int length = 0;
        foreach (string formula in _distinct)
            length += FormulaRenderer.Render(FormulaParser.Parse(formula, _workbook), _workbook).Length;
        return length;
    }
}
