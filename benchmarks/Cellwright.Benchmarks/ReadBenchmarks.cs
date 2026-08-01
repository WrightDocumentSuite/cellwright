using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace Cellwright.Benchmarks;

/// <summary>
/// Reading every cell of a Rows x 20 workbook (produced once by the Cellwright writer):
/// checksum of numbers plus text length keeps the loops honest.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class ReadBenchmarks
{
    private byte[] _file = null!;

    [Params(100_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Cellwright Benchmarks");
        var data = new BenchmarkData(Rows);
        using var stream = new MemoryStream();
        Produce(data, stream).GetAwaiter().GetResult();
        _file = stream.ToArray();
    }

    private static async Task Produce(BenchmarkData data, MemoryStream stream)
    {
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(stream, new XlsxWriterOptions { LeaveOpen = true });
        XlsxSheetWriter sheet = await writer.BeginSheetAsync("Data");
        for (int i = 0; i < data.Rows; i++)
        {
            sheet.WriteRow(
                data.Numbers[i], data.Texts[i], data.Numbers[i] * 2, data.Dates[i], i % 2 == 0,
                data.Numbers[i], data.Texts[i], data.Numbers[i] * 2, data.Dates[i], i % 2 == 0,
                data.Numbers[i], data.Texts[i], data.Numbers[i] * 2, data.Dates[i], i % 2 == 0,
                data.Numbers[i], data.Texts[i], data.Numbers[i] * 2, data.Dates[i], i % 2 == 0);
            await sheet.FlushIfNeededAsync();
        }
    }

    [Benchmark(Baseline = true)]
    public async Task<double> CellwrightStreaming()
    {
        using var stream = new MemoryStream(_file);
        double checksum = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(stream, new XlsxReaderOptions { LeaveOpen = true });
        await foreach (RowData row in reader.ReadRowsAsync(0))
        {
            foreach (ref readonly CellValue cell in row.Cells)
            {
                if (cell.IsNumber)
                    checksum += cell.GetNumber();
                else if (cell.IsText)
                    checksum += cell.GetText().Length;
            }
        }

        return checksum;
    }

    [Benchmark]
    public async Task<double> CellwrightDom()
    {
        using var stream = new MemoryStream(_file);
        double checksum = 0;
        Workbook book = await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true });
        Worksheet sheet = book.Sheets[0];
        RangeRef used = sheet.UsedRange!.Value;
        for (int row = 1; row <= used.End.Row; row++)
        {
            for (int column = 1; column <= used.End.Column; column++)
            {
                CellValue cell = sheet[row, column];
                if (cell.IsNumber)
                    checksum += cell.GetNumber();
                else if (cell.IsText)
                    checksum += cell.GetText().Length;
            }
        }

        return checksum;
    }

    [Benchmark]
    public double ClosedXml()
    {
        using var stream = new MemoryStream(_file);
        double checksum = 0;
        using var workbook = new XLWorkbook(stream);
        IXLWorksheet sheet = workbook.Worksheet(1);
        IXLRange used = sheet.RangeUsed()!;
        foreach (IXLCell cell in used.Cells())
        {
            XLCellValue value = cell.Value;
            if (value.IsNumber)
                checksum += value.GetNumber();
            else if (value.IsText)
                checksum += value.GetText().Length;
        }

        return checksum;
    }

    [Benchmark]
    public double EpPlus()
    {
        using var stream = new MemoryStream(_file);
        double checksum = 0;
        using var package = new ExcelPackage(stream);
        ExcelWorksheet sheet = package.Workbook.Worksheets[0];
        ExcelAddressBase dimension = sheet.Dimension!;
        for (int row = 1; row <= dimension.End.Row; row++)
        {
            for (int column = 1; column <= dimension.End.Column; column++)
            {
                object? value = sheet.Cells[row, column].Value;
                if (value is double number)
                    checksum += number;
                else if (value is string text)
                    checksum += text.Length;
            }
        }

        return checksum;
    }
}
