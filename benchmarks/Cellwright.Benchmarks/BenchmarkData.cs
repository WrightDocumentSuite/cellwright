namespace Cellwright.Benchmarks;

/// <summary>
/// Deterministic dataset shared by all benchmarks: Rows x 20 columns mixing numbers,
/// repeating strings, dates and booleans — a shape typical for export scenarios.
/// </summary>
public sealed class BenchmarkData
{
    public const int Columns = 20;

    public BenchmarkData(int rows)
    {
        Rows = rows;
        var random = new Random(42);
        Numbers = new double[rows];
        Texts = new string[rows];
        Dates = new DateTime[rows];
        for (int i = 0; i < rows; i++)
        {
            Numbers[i] = random.NextDouble() * 10_000;
            Texts[i] = "customer-" + random.Next(500);
            Dates[i] = new DateTime(2024, 1, 1).AddMinutes(random.Next(1_000_000));
        }
    }

    public int Rows { get; }

    public double[] Numbers { get; }

    public string[] Texts { get; }

    public DateTime[] Dates { get; }
}
