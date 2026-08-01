using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace Cellwright.Benchmarks;

/// <summary>
/// Writing Rows x 20 mixed-type cells to an in-memory xlsx: Cellwright streaming,
/// Cellwright DOM, ClosedXML and EPPlus, each through its idiomatic per-cell API.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class WriteBenchmarks
{
    private BenchmarkData _data = null!;

    [Params(100_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _data = new BenchmarkData(Rows);
        ExcelPackage.License.SetNonCommercialPersonal("Cellwright Benchmarks");
    }

    [Benchmark(Baseline = true)]
    public async Task<long> CellwrightStreaming()
    {
        using var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(stream, new XlsxWriterOptions { LeaveOpen = true }))
        {
            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data");
            for (int i = 0; i < _data.Rows; i++)
            {
                WriteCellwrightRow(sheet, i);
                await sheet.FlushIfNeededAsync();
            }
        }

        return stream.Length;
    }

    [Benchmark]
    public async Task<long> CellwrightDom()
    {
        using var stream = new MemoryStream();
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        for (int i = 0; i < _data.Rows; i++)
        {
            int row = i + 1;
            for (int group = 0; group < 4; group++)
            {
                sheet[row, group * 5 + 1] = _data.Numbers[i];
                sheet[row, group * 5 + 2] = _data.Texts[i];
                sheet[row, group * 5 + 3] = _data.Numbers[i] * 2;
                sheet[row, group * 5 + 4] = _data.Dates[i];
                sheet[row, group * 5 + 5] = i % 2 == 0;
            }
        }

        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true });
        return stream.Length;
    }

    [Benchmark]
    public long ClosedXml()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet sheet = workbook.AddWorksheet("Data");
            for (int i = 0; i < _data.Rows; i++)
            {
                int row = i + 1;
                for (int group = 0; group < 4; group++)
                {
                    sheet.Cell(row, group * 5 + 1).Value = _data.Numbers[i];
                    sheet.Cell(row, group * 5 + 2).Value = _data.Texts[i];
                    sheet.Cell(row, group * 5 + 3).Value = _data.Numbers[i] * 2;
                    sheet.Cell(row, group * 5 + 4).Value = _data.Dates[i];
                    sheet.Cell(row, group * 5 + 5).Value = i % 2 == 0;
                }
            }

            workbook.SaveAs(stream);
        }

        return stream.Length;
    }

    [Benchmark]
    public long EpPlus()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            for (int group = 0; group < 4; group++)
                sheet.Column(group * 5 + 4).Style.Numberformat.Format = "m/d/yy h:mm";

            for (int i = 0; i < _data.Rows; i++)
            {
                int row = i + 1;
                for (int group = 0; group < 4; group++)
                {
                    sheet.Cells[row, group * 5 + 1].Value = _data.Numbers[i];
                    sheet.Cells[row, group * 5 + 2].Value = _data.Texts[i];
                    sheet.Cells[row, group * 5 + 3].Value = _data.Numbers[i] * 2;
                    sheet.Cells[row, group * 5 + 4].Value = _data.Dates[i];
                    sheet.Cells[row, group * 5 + 5].Value = i % 2 == 0;
                }
            }

            package.SaveAs(stream);
        }

        return stream.Length;
    }

    private void WriteCellwrightRow(XlsxSheetWriter sheet, int i) =>
        sheet.WriteRow(
            _data.Numbers[i], _data.Texts[i], _data.Numbers[i] * 2, _data.Dates[i], i % 2 == 0,
            _data.Numbers[i], _data.Texts[i], _data.Numbers[i] * 2, _data.Dates[i], i % 2 == 0,
            _data.Numbers[i], _data.Texts[i], _data.Numbers[i] * 2, _data.Dates[i], i % 2 == 0,
            _data.Numbers[i], _data.Texts[i], _data.Numbers[i] * 2, _data.Dates[i], i % 2 == 0);
}
