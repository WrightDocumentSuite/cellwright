using Cellwright.Formulas;

namespace Cellwright.Samples;

/// <summary>
/// Building a small report with formulas, calculating it and reacting to edits.
/// </summary>
internal static class FormulaSample
{
    public static async Task RunAsync(string outputDirectory)
    {
        var book = Workbook.Create();
        Worksheet rates = book.Sheets.Add("Rates");
        Worksheet report = book.Sheets.Add("Report");

        rates["A1"] = "Region";
        rates["B1"] = "Multiplier";
        string[] regions = ["North", "South", "East"];
        for (int i = 0; i < regions.Length; i++)
        {
            rates[i + 2, 1] = regions[i];
            rates[i + 2, 2] = 1 + ((i + 1) * 0.05);
        }

        book.Names.Add("Tax", "0.2");

        report["A1"] = "Region";
        report["B1"] = "Units";
        report["C1"] = "Price";
        report["D1"] = "Net";
        report["E1"] = "Gross";
        for (int row = 2; row <= 4; row++)
        {
            report[row, 1] = regions[row - 2];
            report[row, 2] = row * 10;
            report[row, 3] = 12.5 * row;
            report.Cell(row, 4).SetFormula($"B{row}*C{row}*VLOOKUP(A{row},Rates!$A$2:$B$4,2,FALSE)");
            report.Cell(row, 5).SetFormula($"ROUND(D{row}*(1+Tax),2)");
        }

        report.Cell("D5").SetFormula("SUM(D2:D4)");
        report.Cell("E5").SetFormula("SUM(E2:E4)");
        report.Cell("A6").SetFormula("\"High regions: \"&COUNTIF(E2:E4,\">500\")");

        CalculationSummary first = book.Calculate();
        Console.WriteLine($"  formulas: {first}");
        Console.WriteLine($"  gross total: {report["E5"].GetNumber():F2}");
        Console.WriteLine($"  summary line: {report["A6"].GetText()}");

        // Editing one input recalculates only what depends on it.
        report["B2"] = 500;
        CalculationSummary second = book.Calculate();
        Console.WriteLine($"  after edit: {second}, gross total {report["E5"].GetNumber():F2}");

        // Ad-hoc evaluation without storing a formula anywhere.
        Console.WriteLine($"  ad-hoc average: {book.Evaluate("AVERAGE(Report!D2:D4)").GetNumber():F2}");

        await book.SaveAsync(Path.Combine(outputDirectory, "formulas.xlsx"));
        Console.WriteLine("  formulas.xlsx written with cached values (Excel opens it without recalculating)");
    }
}
