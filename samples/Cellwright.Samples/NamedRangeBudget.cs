using Cellwright;
using Cellwright.Formulas;

namespace Cellwright.Samples;

/// <summary>Uses an editable named rate, input validation and recalculation for a small budget.</summary>
internal static class NamedRangeBudget
{
    public static async Task RunAsync(string directory)
    {
        Workbook book = Workbook.Create();
        Worksheet settings = book.Sheets.Add("Settings");
        settings["A1"] = "Tax rate";
        settings["B1"] = 0.2;
        book.Names.Add("TaxRate", "Settings!$B$1");

        Worksheet budget = book.Sheets.Add("Budget");
        budget["A1"] = "Services";
        budget["B1"] = 1000;
        budget["A2"] = "Gross total";
        budget.Cell("B2").SetFormula("B1*(1+TaxRate)");
        budget.Range("B1:B2").NumberFormat("#,##0.00");
        budget.Validations.Add(DataValidation.WholeBetween("B1", 0, 1000000));
        book.Calculate();
        double initial = budget["B2"].GetNumber();
        budget["B1"] = 1500;
        book.Calculate();
        double updated = budget["B2"].GetNumber();
        if (Math.Abs(initial - 1200) > 0.001 || Math.Abs(updated - 1800) > 0.001)
            throw new InvalidOperationException("The named-rate formula did not recalculate.");

        await book.SaveAsync(Path.Combine(directory, "named-range-budget.xlsx"));
        Console.WriteLine($"  Named-rate budget: {initial} -> {updated}, cached results saved");
    }
}
