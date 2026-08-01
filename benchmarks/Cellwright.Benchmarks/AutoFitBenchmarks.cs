using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Pdf;
using Cellwright.Styles;

namespace Cellwright.Benchmarks;

/// <summary>
/// Sizing columns and rows to their contents. The cost is the same measurement a render pays —
/// format the value, measure the string against the font — but paid once per cell rather than
/// once per cell per page, and with no output to write afterwards.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class AutoFitBenchmarks
{
    private Workbook _book = null!;

    /// <summary>Rows of the sample sheet; five columns each.</summary>
    [Params(20_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup() => _book = Build();

    [Benchmark(Baseline = true)]
    public double? Columns()
    {
        Worksheet sheet = _book.Sheets[0];
        sheet.AutoFitColumns();
        return sheet.Column(1).Width;
    }

    /// <summary>The first few hundred rows usually decide the answer; this is what capping costs.</summary>
    [Benchmark]
    public double? ColumnsSampled()
    {
        Worksheet sheet = _book.Sheets[0];
        sheet.AutoFitColumns(new AutoFitOptions { SampleRows = 500 });
        return sheet.Column(1).Width;
    }

    [Benchmark]
    public double? ColumnsAndRows()
    {
        Worksheet sheet = _book.Sheets[0];
        sheet.AutoFit();
        return sheet.Row(1).Height;
    }

    private Workbook Build()
    {
        var data = new BenchmarkData(Rows);
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");

        StyleId money = book.RegisterStyle(CellStyle.Default with { NumberFormat = "#,##0.00" });
        StyleId date = book.RegisterStyle(CellStyle.Default with { NumberFormat = "yyyy-mm-dd" });

        for (int i = 0; i < data.Rows; i++)
        {
            int row = i + 1;
            sheet[row, 1] = CellValue.Text(data.Texts[i]);
            sheet[row, 2] = CellValue.Number(data.Numbers[i]);
            sheet[row, 3] = CellValue.Number(data.Numbers[i] * 2);
            sheet[row, 4] = CellValue.Number(45000 + (i % 900));
            sheet[row, 5] = CellValue.Boolean(i % 2 == 0);

            sheet.Cell(new CellRef(row, 2)).SetStyle(money);
            sheet.Cell(new CellRef(row, 3)).SetStyle(money);
            sheet.Cell(new CellRef(row, 4)).SetStyle(date);
        }

        return book;
    }
}
