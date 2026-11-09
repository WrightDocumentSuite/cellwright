using Cellwright;

namespace Cellwright.Samples;

/// <summary>Streams quoted CSV fields, converts to XLSX and exports the worksheet back to CSV.</summary>
internal static class CsvRoundTrip
{
    public static async Task RunAsync(string directory)
    {
        string source = Path.Combine(directory, "sales-input.csv");
        await using (CsvWriter writer = CsvWriter.Create(source))
        {
            await writer.WriteRowAsync(new[] { CellValue.Text("Product"), CellValue.Text("Amount") });
            await writer.WriteRowAsync(new[] { CellValue.Text("Starter, boxed"), CellValue.Number(1250.5) });
            await writer.WriteRowAsync(new[] { CellValue.Text("Support \"Plus\""), CellValue.Number(320) });
        }

        Workbook book = await Workbook.LoadCsvAsync(source);
        await book.SaveAsync(Path.Combine(directory, "sales-from-csv.xlsx"));
        string exported = Path.Combine(directory, "sales-export.csv");
        await book.Sheets[0].SaveCsvAsync(exported);
        int rows = 0;
        await foreach (RowData row in CsvReader.ReadRowsAsync(exported))
        {
            if (rows == 1 && row.GetValueOrBlank(0).GetText() != "Starter, boxed")
                throw new InvalidOperationException("The CSV field containing a comma did not round-trip.");
            rows++;
        }

        if (rows != 3)
            throw new InvalidOperationException("The CSV row count changed.");
        Console.WriteLine($"  CSV round-trip: {rows} rows, including quoted commas and quotes");
    }
}
