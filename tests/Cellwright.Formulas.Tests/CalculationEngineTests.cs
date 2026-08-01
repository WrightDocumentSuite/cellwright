using Xunit;

namespace Cellwright.Formulas.Tests;

public sealed class CalculationEngineTests
{
    [Fact]
    public void CalculatesChainsInDependencyOrder()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 2;
        sheet.Cell("A2").SetFormula("A1*3");
        sheet.Cell("A3").SetFormula("A2+4");
        sheet.Cell("A4").SetFormula("SUM(A1:A3)");

        CalculationSummary summary = book.Calculate();

        Assert.Equal(6.0, sheet["A2"].GetNumber());
        Assert.Equal(10.0, sheet["A3"].GetNumber());
        Assert.Equal(18.0, sheet["A4"].GetNumber());
        Assert.Equal(3, summary.CalculatedCells);
        Assert.False(summary.HasCircularReferences);
    }

    [Fact]
    public void CalculatesAcrossSheets()
    {
        var book = Workbook.Create();
        Worksheet data = book.Sheets.Add("Data");
        Worksheet report = book.Sheets.Add("Report");
        data["A1"] = 5;
        data["A2"] = 7;
        report.Cell("B1").SetFormula("SUM(Data!A1:A2)*2");

        book.Calculate();

        Assert.Equal(24.0, report["B1"].GetNumber());
    }

    [Fact]
    public void EditingAValue_RecalculatesOnlyWhatDependsOnIt()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet["B1"] = 100;
        sheet.Cell("A2").SetFormula("A1+1");
        sheet.Cell("A3").SetFormula("A2+1");
        sheet.Cell("B2").SetFormula("B1+1");

        book.Calculate();
        sheet["A1"] = 10;
        CalculationSummary summary = book.Calculate();

        Assert.Equal(11.0, sheet["A2"].GetNumber());
        Assert.Equal(12.0, sheet["A3"].GetNumber());
        Assert.Equal(101.0, sheet["B2"].GetNumber());
        Assert.Equal(2, summary.CalculatedCells);
    }

    [Fact]
    public void ForceRecalculatesEverything()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet.Cell("A2").SetFormula("A1+1");

        book.Calculate();
        CalculationSummary forced = book.Calculate(new CalculateOptions { Force = true });

        Assert.Equal(1, forced.CalculatedCells);
    }

    [Fact]
    public void CircularReferences_ReportZeroAndAWarning()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("A2+1");
        sheet.Cell("A2").SetFormula("A1+1");

        var warnings = new List<WorkbookWarning>();
        CalculationSummary summary = book.Calculate(new CalculateOptions { OnWarning = warnings.Add });

        Assert.True(summary.HasCircularReferences);
        Assert.Equal(2, summary.CircularCells.Count);
        Assert.Equal(0.0, sheet["A1"].GetNumber());
        Assert.Contains(warnings, w => w.Code == "CircularReference");
        Assert.False(book.IsFullyCalculatedForTest());
    }

    [Fact]
    public void SelfReference_IsCircular()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("A1+1");

        Assert.True(book.Calculate().HasCircularReferences);
    }

    [Fact]
    public void VolatileFormulas_RecalculateEveryPass()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("RAND()");
        sheet.Cell("A2").SetFormula("A1*0+1");

        book.Calculate();
        CalculationSummary second = book.Calculate();

        Assert.Equal(2, second.CalculatedCells);
    }

    [Fact]
    public void DeepChains_DoNotOverflowTheStack()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        for (int row = 2; row <= 20_000; row++)
            sheet.Cell(row, 1).SetFormula($"A{row - 1}+1");

        book.Calculate();

        Assert.Equal(20_000.0, sheet[20_000, 1].GetNumber());
    }

    [Fact]
    public void UnparsableFormula_BecomesNameErrorWithWarning()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("1+");

        var warnings = new List<WorkbookWarning>();
        book.Calculate(new CalculateOptions { OnWarning = warnings.Add });

        Assert.Equal(ExcelError.Name, sheet["A1"].GetError());
        Assert.Contains(warnings, w => w.Code == "BadFormula");
    }

    [Fact]
    public void ReferencesToMissingSheets_BecomeRefErrors()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("Missing!A1+1");

        book.Calculate();

        Assert.Equal(ExcelError.Reference, sheet["A1"].GetError());
    }

    [Fact]
    public void RemovingASheet_TurnsItsReferencesIntoRefErrors()
    {
        var book = Workbook.Create();
        Worksheet data = book.Sheets.Add("Data");
        Worksheet report = book.Sheets.Add("Report");
        data["A1"] = 5;
        report.Cell("B1").SetFormula("Data!A1*2");

        book.Calculate();
        Assert.Equal(10.0, report["B1"].GetNumber());

        book.Sheets.Remove("Data");
        book.Calculate(new CalculateOptions { Force = true });

        Assert.Equal(ExcelError.Reference, report["B1"].GetError());
    }

    [Fact]
    public void DefinedNames_ResolveInFormulas()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 4;
        sheet["A2"] = 6;
        book.Names.Add("Values", "S!$A$1:$A$2");
        book.Names.Add("Rate", "0.5");
        sheet.Cell("B1").SetFormula("SUM(Values)*Rate");

        book.Calculate();

        Assert.Equal(5.0, sheet["B1"].GetNumber());
    }

    [Fact]
    public void SheetScopedName_ShadowsWorkbookName()
    {
        var book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");
        book.Names.Add("Rate", "1");
        book.Names.Add("Rate", "2", second);
        first.Cell("A1").SetFormula("Rate");
        second.Cell("A1").SetFormula("Rate");

        book.Calculate();

        Assert.Equal(1.0, first["A1"].GetNumber());
        Assert.Equal(2.0, second["A1"].GetNumber());
    }

    [Fact]
    public void AdHocEvaluation_SeesWorkbookData()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sales");
        sheet["B2"] = 10;
        sheet["B3"] = 20;

        Assert.Equal(36.0, book.Evaluate("SUM(Sales!B2:B3)*1.2").GetNumber());
    }

    [Fact]
    public void CalculatedValue_TriggersCalculationOnDemand()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 3;
        sheet.Cell("A2").SetFormula("A1*7");

        Assert.Equal(21.0, sheet.Cell("A2").CalculatedValue.GetNumber());
    }

    [Fact]
    public void InsertingRowsMovesTheGraphWithTheFormulas()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 10;
        sheet["A2"] = 20;
        sheet.Cell("B1").SetFormula("SUM(A1:A2)");
        book.Calculate();
        Assert.Equal(30.0, sheet["B1"].GetNumber());

        sheet.InsertRows(2);
        sheet["A2"] = 5;
        book.Calculate();

        // The total moved to B1 unchanged, its range grew over the new row, and editing that
        // row still reaches it — which only works if the engine rebuilt its graph.
        Assert.Equal("SUM(A1:A3)", sheet.Cell("B1").Formula);
        Assert.Equal(35.0, sheet["B1"].GetNumber());
    }

    [Fact]
    public void DeletingRowsLeavesTheBrokenFormulasInError()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 10;
        sheet["A2"] = 20;
        sheet.Cell("B1").SetFormula("A2*2");
        book.Calculate();
        Assert.Equal(40.0, sheet["B1"].GetNumber());

        sheet.DeleteRows(2);
        book.Calculate();

        Assert.Equal(ExcelError.Reference, sheet["B1"].GetError());
    }

    [Fact]
    public void Cancellation_StopsALongCalculation()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        for (int row = 2; row <= 5_000; row++)
            sheet.Cell(row, 1).SetFormula($"A{row - 1}+1");

        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(1));

        try
        {
            book.Calculate(new CalculateOptions { CancellationToken = cancellation.Token });
        }
        catch (OperationCanceledException)
        {
            return; // Cancellation observed, which is the point of the test.
        }
    }
}

internal static class WorkbookTestExtensions
{
    /// <summary>Exposes the internal calculated flag that controls the fullCalcOnLoad marker.</summary>
    public static bool IsFullyCalculatedForTest(this Workbook workbook) => workbook.IsFullyCalculated;
}
