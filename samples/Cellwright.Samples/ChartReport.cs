using Cellwright;

namespace Cellwright.Samples;

/// <summary>Builds a native column chart linked to worksheet data and checks its saved series.</summary>
internal static class ChartReport
{
    public static async Task RunAsync(string directory)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sales");
        sheet["A1"] = "Month";
        sheet["B1"] = "Revenue";
        string[] months = ["January", "February", "March"];
        double[] revenue = [12000, 15500, 14250];
        for (int i = 0; i < months.Length; i++)
        {
            sheet[i + 2, 1] = months[i];
            sheet[i + 2, 2] = revenue[i];
        }

        sheet.Range("A1:B1").Style(style => style.Bold());
        sheet.Range("B2:B4").NumberFormat("#,##0.00");
        sheet.AddChart(ChartKind.Column, "D2:L18", "Sales!$B$2:$B$4",
            categoriesReference: "Sales!$A$2:$A$4", title: "Quarterly revenue", seriesName: "Revenue");
        string path = Path.Combine(directory, "chart-report.xlsx");
        await book.SaveAsync(path);

        Workbook reloaded = await Workbook.LoadAsync(path);
        Chart chart = reloaded.Sheets[0].Charts.Single();
        if (chart.Title != "Quarterly revenue" || !chart.Series.Single().Values.SequenceEqual(revenue.Select(value => (double?)value)))
            throw new InvalidOperationException("The saved chart no longer matches its worksheet data.");
        Console.WriteLine($"  Chart report: {chart.Series.Single().Values.Count} plotted values verified");
    }
}
