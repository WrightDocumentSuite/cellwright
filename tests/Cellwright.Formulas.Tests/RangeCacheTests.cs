using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The range cache trades a copy per lookup for a shared array, so the tests that matter are
/// the ones proving nobody can observe a stale or mutated copy.
/// </summary>
public sealed class RangeCacheTests
{
    private static Workbook BuildLookupWorkbook(string tableSheetName, int rows = 200)
    {
        var book = Workbook.Create();
        Worksheet table = book.Sheets.Add(tableSheetName);
        for (int row = 1; row <= rows; row++)
        {
            table[row, 1] = row;
            table[row, 2] = row * 10;
        }

        Worksheet report = tableSheetName == "Data" ? book.Sheets.Add("Report") : table;
        int firstFormulaRow = ReferenceEquals(report, table) ? rows + 2 : 1;
        for (int i = 0; i < 50; i++)
        {
            int row = firstFormulaRow + i;
            report[row, 4] = (i % rows) + 1;
            report.Cell(row, 5).SetFormula($"VLOOKUP(D{row},{Quote(tableSheetName)}!$A$1:$B${rows},2,FALSE)");
        }

        return book;
    }

    private static string Quote(string sheetName) => sheetName;

    [Fact]
    public void RepeatedLookupsAgainstOneTableMaterializeItOnce()
    {
        Workbook book = BuildLookupWorkbook("Data");
        RangeCache cache = CalculationExtensions.RangeCacheOf(book);

        book.Calculate();

        Assert.Equal(1, cache.Misses);
        Assert.Equal(49, cache.Hits);
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public void LookupTableOnTheSameSheetStillBenefits()
    {
        // The case that rules out a per-sheet version counter: results are written to the very
        // sheet holding the table, so a coarse invalidation would miss on every single call.
        Workbook book = BuildLookupWorkbook("Everything");
        RangeCache cache = CalculationExtensions.RangeCacheOf(book);

        book.Calculate();

        Assert.Equal(1, cache.Misses);
        Assert.True(cache.Hits >= 49, $"expected the cache to serve the repeated lookups, saw {cache.Hits} hit(s)");
        Worksheet sheet = book.Sheets["Everything"];
        Assert.Equal(70.0, sheet[sheet.UsedRange!.Value.End.Row - 43, 5].GetNumber());
    }

    [Fact]
    public void EditingTheTableChangesLaterResults()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= 100; row++)
        {
            sheet[row, 1] = row;
            sheet[row, 2] = row * 10;
        }

        sheet.Cell("D1").SetFormula("VLOOKUP(7,$A$1:$B$100,2,FALSE)");
        book.Calculate();
        Assert.Equal(70.0, sheet["D1"].GetNumber());

        sheet[7, 2] = 999;
        book.Calculate();

        Assert.Equal(999.0, sheet["D1"].GetNumber());
    }

    [Fact]
    public void ComputedValuesInsideARangeInvalidateIt()
    {
        // The table is itself produced by formulas, so the engine writes into the very
        // rectangle other formulas read.
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= 100; row++)
        {
            sheet[row, 1] = row;
            sheet.Cell(row, 2).SetFormula($"A{row}*10");
        }

        sheet.Cell("D1").SetFormula("SUM($B$1:$B$100)");
        book.Calculate();

        Assert.Equal(50_500.0, sheet["D1"].GetNumber());

        sheet[7, 1] = 1_000;
        book.Calculate();
        Assert.Equal(50_500.0 - 70 + 10_000, sheet["D1"].GetNumber());
    }

    [Fact]
    public void VolatileReadOfAnUncalculatedRangeDoesNotPoisonTheCache()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= 60; row++)
        {
            sheet[row, 1] = row;
            sheet.Cell(row, 2).SetFormula($"A{row}+1");
        }

        // OFFSET is volatile and creates a dependency the static graph cannot see, so it may
        // well read column B before those formulas run.
        sheet.Cell("D1").SetFormula("SUM(OFFSET($B$1,0,0,60,1))");
        sheet.Cell("D2").SetFormula("SUM($B$1:$B$60)");

        book.Calculate();
        book.Calculate(new CalculateOptions { Force = true });

        Assert.Equal(1_890.0, sheet["D2"].GetNumber());
        Assert.Equal(sheet["D2"].GetNumber(), sheet["D1"].GetNumber());
    }

    [Fact]
    public void SmallRangesAreNotCached()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        sheet["A2"] = 2;
        sheet.Cell("C1").SetFormula("SUMPRODUCT($A$1:$A$2,$A$1:$A$2)");
        RangeCache cache = CalculationExtensions.RangeCacheOf(book);

        book.Calculate();

        Assert.Equal(0, cache.Count);
        Assert.Equal(5.0, sheet["C1"].GetNumber());
    }

    [Fact]
    public void TheCacheStaysWithinItsBudget()
    {
        var cache = new RangeCache();
        const int Cells = 300_000;
        int ranges = (int)((RangeCache.MaxCells / Cells) + 3);

        for (int i = 0; i < ranges; i++)
        {
            var key = new RangeKey(0, (i * 1000) + 1, 1, (i * 1000) + Cells, 1);
            cache.Add(new CachedRange(key, new ArrayValue(Cells, 1, new CellValue[Cells])));
        }

        Assert.True(cache.RetainedCells <= RangeCache.MaxCells,
            $"retained {cache.RetainedCells} cells, budget is {RangeCache.MaxCells}");
        Assert.True(cache.Count < ranges, "nothing was evicted");
    }

    [Fact]
    public void EvictionKeepsResultsCorrect()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= 500; row++)
            sheet[row, 1] = row;

        // Distinct overlapping ranges force repeated materialization and eviction pressure.
        for (int i = 0; i < 40; i++)
            sheet.Cell(i + 1, 3).SetFormula($"SUM($A$1:$A${400 + i})");

        book.Calculate();

        for (int i = 0; i < 40; i++)
        {
            int last = 400 + i;
            double expected = (double)last * (last + 1) / 2;
            Assert.Equal(expected, sheet[i + 1, 3].GetNumber());
        }
    }

    [Fact]
    public void FunctionsDoNotMutateTheSharedArray()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= 50; row++)
        {
            sheet[row, 1] = row;
            sheet[row, 2] = row % 5;
        }

        // Every one of these reads the same rectangle through a different code path.
        sheet.Cell("D1").SetFormula("SUM($A$1:$A$50)");
        sheet.Cell("D2").SetFormula("SUMIF($A$1:$A$50,\">25\")");
        sheet.Cell("D3").SetFormula("SUMPRODUCT($A$1:$A$50,$A$1:$A$50)");
        sheet.Cell("D4").SetFormula("MATCH(33,$A$1:$A$50,0)");
        sheet.Cell("D5").SetFormula("SUM($A$1:$A$50*2)");
        book.Calculate();

        Assert.Equal(1_275.0, sheet["D1"].GetNumber());
        Assert.Equal(950.0, sheet["D2"].GetNumber());
        Assert.Equal(42_925.0, sheet["D3"].GetNumber());
        Assert.Equal(33.0, sheet["D4"].GetNumber());
        Assert.Equal(2_550.0, sheet["D5"].GetNumber());

        // A second pass must see the same numbers: a mutated shared array would show up here.
        book.Calculate(new CalculateOptions { Force = true });
        Assert.Equal(1_275.0, sheet["D1"].GetNumber());
        Assert.Equal(42_925.0, sheet["D3"].GetNumber());
    }
}
