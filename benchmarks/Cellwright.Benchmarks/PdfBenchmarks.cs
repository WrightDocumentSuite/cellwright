using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Pdf;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Benchmarks;

/// <summary>
/// Rendering a sheet to PDF. The cost is measurement rather than output: every cell's text is
/// measured against real font metrics, and a column of numbers is formatted before it can be.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class PdfBenchmarks
{
    private Workbook _plain = null!;
    private Workbook _styled = null!;
    private Workbook _conditional = null!;
    private Workbook _graded = null!;

    /// <summary>Rows of the sample sheet; five columns each.</summary>
    [Params(20_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _plain = Build(styled: false);
        _styled = Build(styled: true);

        // A rule per column of numbers, the shape a report actually carries.
        _conditional = Build(styled: true);
        Worksheet rules = _conditional.Sheets[0];
        rules.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            $"B1:B{Rows}", 5000, DifferentialStyle.Bad));
        rules.ConditionalFormats.Add(ConditionalFormatRule.CellLessThan(
            $"C1:C{Rows}", 5000, DifferentialStyle.Good));

        // The graded rules are the ones that have to look at the whole range before they can
        // judge any one cell.
        _graded = Build(styled: true);
        _graded.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.Scale($"B1:B{Rows}", ColorScale.RedYellowGreen));
        _graded.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.Bars($"C1:C{Rows}", new DataBar()));
    }

    [Benchmark(Baseline = true)]
    public int Plain() => Render(_plain, tagged: false);

    [Benchmark]
    public int Styled() => Render(_styled, tagged: false);

    [Benchmark]
    public int Tagged() => Render(_styled, tagged: true);

    /// <summary>What working out two comparison rules over the sheet costs.</summary>
    [Benchmark]
    public int Conditional() => Render(_conditional, tagged: false);

    /// <summary>The same for a colour scale and a data bar, which measure the whole range first.</summary>
    [Benchmark]
    public int Graded() => Render(_graded, tagged: false);

    private static int Render(Workbook book, bool tagged)
    {
        PdfExportResult result = PdfExporter.Render(book, new PdfExportOptions { Tagged = tagged });
        using PdfDocument pdf = result.Document;
        return result.PageCount;
    }

    private Workbook Build(bool styled)
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

            if (!styled)
                continue;

            sheet.Cell(new CellRef(row, 2)).SetStyle(money);
            sheet.Cell(new CellRef(row, 3)).SetStyle(money);
            sheet.Cell(new CellRef(row, 4)).SetStyle(date);
        }

        return book;
    }
}
