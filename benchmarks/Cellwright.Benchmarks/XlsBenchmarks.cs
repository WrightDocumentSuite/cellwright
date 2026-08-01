using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using Cellwright.Xls;
using ExcelDataReader;

namespace Cellwright.Benchmarks;

/// <summary>
/// Reading and writing the legacy binary format, against ExcelDataReader over the same bytes.
/// The workbook is produced once by this library's own writer so that both readers see exactly
/// the same file.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, launchCount: 1, warmupCount: 1, iterationCount: 3)]
public class XlsBenchmarks
{
    private byte[] _file = null!;
    private Workbook _workbook = null!;

    /// <summary>Rows of the sample; the old grid stops at 65536, so this stays under it.</summary>
    [Params(50_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var data = new BenchmarkData(Rows);
        _workbook = Workbook.Create();
        Worksheet sheet = _workbook.Sheets.Add("Data");
        for (int i = 0; i < data.Rows; i++)
        {
            int row = i + 1;
            sheet[row, 1] = CellValue.Number(data.Numbers[i]);
            sheet[row, 2] = CellValue.Text(data.Texts[i]);
            sheet[row, 3] = CellValue.Number(data.Numbers[i] * 2);
            sheet[row, 4] = CellValue.Boolean(i % 2 == 0);
        }

        using var stream = new MemoryStream();
        XlsWorkbook.SaveAsync(_workbook, stream, new XlsSaveOptions { LeaveOpen = true })
            .AsTask().GetAwaiter().GetResult();
        _file = stream.ToArray();
    }

    [Benchmark(Baseline = true)]
    public async Task<double> CellwrightRead()
    {
        using var stream = new MemoryStream(_file);
        Workbook workbook = await XlsWorkbook.LoadAsync(stream, new XlsLoadOptions { LeaveOpen = true });
        Worksheet sheet = workbook.Sheets[0];

        double checksum = 0;
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

    /// <summary>The same read without a model: rows arrive one at a time and nothing is kept.</summary>
    [Benchmark]
    public async Task<double> CellwrightStreamRows()
    {
        using var stream = new MemoryStream(_file);
        await using XlsReader reader = await XlsReader.OpenAsync(stream, new XlsLoadOptions { LeaveOpen = true });

        double checksum = 0;
        await foreach (RowData row in reader.ReadRowsAsync(0))
        {
            foreach (CellValue cell in row.Cells)
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
    public double ExcelDataReaderRead()
    {
        using var stream = new MemoryStream(_file);
        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

        double checksum = 0;
        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                object? value = reader.GetValue(i);
                if (value is double number)
                    checksum += number;
                else if (value is string text)
                    checksum += text.Length;
            }
        }

        return checksum;
    }

    [Benchmark]
    public async Task<long> CellwrightWrite()
    {
        using var stream = new MemoryStream(_file.Length);
        await XlsWorkbook.SaveAsync(_workbook, stream, new XlsSaveOptions { LeaveOpen = true });
        return stream.Length;
    }
}
