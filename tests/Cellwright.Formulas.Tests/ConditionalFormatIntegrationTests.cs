using Xunit;

namespace Cellwright.Formulas.Tests;

public sealed class ConditionalFormatIntegrationTests
{
    [Fact]
    public void RelativeConditionalFormulaEvaluatesAgainstEachTargetCell()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        sheet["A1"] = 1;
        sheet["A2"] = 0;
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression("B1:B2", "$A1>0", DifferentialStyle.Good));

        var evaluator = new ConditionalFormatEvaluator(
            sheet,
            evaluateRelative: (formula, anchor, cell) =>
                sheet.EvaluateCopied(formula, anchor, cell));

        Assert.Equal(DifferentialStyle.Good, evaluator.Of(CellRef.Parse("B1")).Style);
        Assert.Null(evaluator.Of(CellRef.Parse("B2")).Style);
        Assert.Empty(evaluator.Unsupported);
    }

    [Fact]
    public void ARelativeRuleKeepsOneCompiledProgramAcrossManyTargets()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        const int Rows = 2_000;
        for (int row = 1; row <= Rows; row++)
            sheet[row, 1] = row % 2;

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression($"B1:B{Rows}", "$A1>0", DifferentialStyle.Good));
        var evaluator = new ConditionalFormatEvaluator(
            sheet,
            evaluateRelative: (formula, anchor, cell) =>
                sheet.EvaluateCopied(formula, anchor, cell));

        for (int row = 1; row <= Rows; row++)
        {
            DifferentialStyle? expected = row % 2 == 1 ? DifferentialStyle.Good : null;
            Assert.Equal(expected, evaluator.Of(new CellRef(row, 2)).Style);
        }

        Assert.Equal(1, CalculationExtensions.ProgramCountOf(book));
    }

    [Fact]
    public void CopiedSlidingRangesStayDistinctWhileAnAbsoluteRangeIsShared()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        const int Targets = 96;
        for (int row = 1; row <= Targets + 31; row++)
        {
            sheet[row, 1] = row;
            sheet[row, 2] = row * 10;
        }

        RangeCache cache = CalculationExtensions.RangeCacheOf(book);
        var source = new CellRef(1, 4);
        for (int row = 1; row <= Targets; row++)
        {
            CellValue result = sheet.EvaluateCopied(
                "VLOOKUP(A1,A1:B32,2,FALSE)", source, new CellRef(row, 4));
            Assert.Equal(row * 10.0, result.GetNumber());
        }

        Assert.Equal(Targets, cache.Misses);
        Assert.Equal(0, cache.Hits);
        Assert.Equal(Targets, cache.Count);

        cache.Clear();
        cache.ResetCounters();
        for (int row = 1; row <= Targets; row++)
        {
            CellValue result = sheet.EvaluateCopied(
                "VLOOKUP(1,$A$1:$B$32,2,FALSE)", source, new CellRef(row, 4));
            Assert.Equal(10.0, result.GetNumber());
        }

        Assert.Equal(1, cache.Misses);
        Assert.Equal(Targets - 1, cache.Hits);
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public void EvaluateCopiedRejectsCellsOutsideTheWorksheetGrid()
    {
        Worksheet sheet = Workbook.Create().Sheets.Add("Report");

        ArgumentOutOfRangeException source = Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet.EvaluateCopied("ROW()", default, new CellRef(1, 1)));
        ArgumentOutOfRangeException target = Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet.EvaluateCopied(
                "ROW()",
                new CellRef(1, 1),
                new CellRef(SheetLimits.MaxRow + 1, SheetLimits.MaxColumn + 1)));

        Assert.Equal("sourceCell", source.ParamName);
        Assert.Equal("targetCell", target.ParamName);
    }

    [Fact]
    public void DynamicIndirectProgramsStayBoundedAcrossCopiedTargets()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        int targets = EvalContext.MaxDynamicPrograms + 500;
        for (int row = 1; row <= targets; row++)
            sheet[row, 1] = row;

        for (int row = 1; row <= targets; row++)
        {
            CellValue value = sheet.EvaluateCopied(
                "INDIRECT(\"A\"&ROW())",
                new CellRef(1, 2),
                new CellRef(row, 2));
            Assert.Equal(row, value.GetNumber());
        }

        Assert.Equal(EvalContext.MaxDynamicPrograms + 1, CalculationExtensions.ProgramCountOf(book));
    }

    [Fact]
    public void EvaluateCopiedMovesOnlyRelativeReferenceAxes()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        sheet["C4"] = 1;
        sheet["A4"] = 2;
        sheet["C1"] = 4;
        sheet["A1"] = 8;

        CellValue value = sheet.EvaluateCopied(
            "A1+$A1+A$1+$A$1",
            new CellRef(2, 2),
            new CellRef(5, 4));

        Assert.Equal(15.0, value.GetNumber());
        Assert.Equal(1, CalculationExtensions.ProgramCountOf(book));
    }

    [Fact]
    public void EvaluateCopiedReturnsRefWhenACopyLeavesTheGrid()
    {
        Worksheet sheet = Workbook.Create().Sheets.Add("Report");

        CellValue value = sheet.EvaluateCopied(
            "A1",
            new CellRef(2, 2),
            new CellRef(1, 1));

        Assert.Equal(ExcelError.Reference, value.GetError());
    }

    [Fact]
    public void CopiedStructuredThisRowReferenceUsesTheTargetRow()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        sheet["A1"] = "Item";
        sheet["B1"] = "Qty";
        sheet["A2"] = "First";
        sheet["B2"] = 10;
        sheet["A3"] = "Second";
        sheet["B3"] = 20;
        sheet["A4"] = "Total";
        sheet["B4"] = 30;
        sheet.Tables.Add(new ExcelTable
        {
            Name = "Sales",
            Range = RangeRef.Parse("A1:B4"),
            Columns = ["Item", "Qty"],
            HasTotalsRow = true,
        });

        Assert.Equal(20.0, sheet.EvaluateCopied(
            "Sales[@Qty]",
            new CellRef(2, 4),
            new CellRef(3, 4)).GetNumber());
        Assert.Equal(20.0, sheet.EvaluateCopied(
            "[@Qty]",
            new CellRef(2, 2),
            new CellRef(3, 2)).GetNumber());
    }
}
