using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// What a sheet's conditional formatting rules actually come to for a cell: which ones match,
/// in what order they paint over each other, and what the graded ones work out to.
/// </summary>
public sealed class ConditionalFormatEvaluatorTests
{
    private static readonly DateOnly Today = new(2026, 7, 30);

    private static Worksheet Column(params double[] values)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int i = 0; i < values.Length; i++)
            sheet[i + 1, 1] = values[i];

        return sheet;
    }

    private static ConditionalFormatEvaluator Evaluator(
        Worksheet sheet, Func<string, CellRef, CellValue>? evaluate = null) =>
        new(sheet, evaluate, Today);

    private static CellRef At(int row) => new(row, 1);

    [Fact]
    public void AComparisonAgainstALiteralPaintsTheCellsThatPassIt()
    {
        Worksheet sheet = Column(5, 50, 500);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A3", 100, DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Equal(DifferentialStyle.Bad, evaluator.Of(At(3)).Style);
    }

    [Fact]
    public void BetweenIsInclusiveAtBothEnds()
    {
        Worksheet sheet = Column(9, 10, 15, 20, 21);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellBetween("A1:A5", 10, 20, DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.NotNull(evaluator.Of(At(2)).Style);
        Assert.NotNull(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
        Assert.Null(evaluator.Of(At(5)).Style);
    }

    [Fact]
    public void ALaterRulePaintsOverAnEarlierOneOnlyWhereItSaysSomething()
    {
        Worksheet sheet = Column(500);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            Style = new DifferentialStyle { Bold = true, FontColor = ExcelColor.FromRgb(1, 2, 3) },
        });

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["2"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(9, 9, 9) },
        });

        DifferentialStyle? style = Evaluator(sheet).Of(At(1)).Style;

        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.Equal(ExcelColor.FromRgb(1, 2, 3), style.FontColor);
        Assert.Equal(ExcelColor.FromRgb(9, 9, 9), style.FillColor);
    }

    [Fact]
    public void StopIfTrueKeepsLaterRulesFromPainting()
    {
        Worksheet sheet = Column(500);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            StopIfTrue = true,
            Style = new DifferentialStyle { Bold = true },
        });

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["2"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(9, 9, 9) },
        });

        DifferentialStyle? style = Evaluator(sheet).Of(At(1)).Style;

        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.True(style.FillColor.IsAuto);
    }

    [Fact]
    public void TextRulesLookAtWhatTheCellSays()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Overdue invoice";
        sheet["A2"] = "Paid";

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Containing("A1:A2", "overdue", DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
    }

    [Fact]
    public void DuplicatesAndUniquesAreCountedAcrossTheWholeRange()
    {
        Worksheet sheet = Column(1, 2, 2, 3);

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Duplicates("A1:A4", DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.NotNull(evaluator.Of(At(2)).Style);
        Assert.NotNull(evaluator.Of(At(3)).Style);
        Assert.Null(evaluator.Of(At(4)).Style);
    }

    [Fact]
    public void TheTopValuesAreTheHighestOnesInTheRange()
    {
        Worksheet sheet = Column(10, 20, 30, 40, 50);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Top("A1:A5", 2, DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
        Assert.NotNull(evaluator.Of(At(5)).Style);
    }

    [Fact]
    public void AboveAverageComparesAgainstTheRangesMean()
    {
        Worksheet sheet = Column(1, 2, 3, 4, 5);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A5")],
            Type = ConditionalFormatType.AboveAverage,
            Above = true,
            Style = DifferentialStyle.Good,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        // The mean is 3, and "above" is strict.
        Assert.Null(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
    }

    [Fact]
    public void AColourScaleGradesBetweenItsStops()
    {
        Worksheet sheet = Column(0, 50, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A3", new ColorScale
        {
            Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
            Colors = [ExcelColor.FromRgb(0, 0, 0), ExcelColor.FromRgb(0xFF, 0xFF, 0xFF)],
        }));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(ExcelColor.FromRgb(0, 0, 0), evaluator.Of(At(1)).ScaleColor);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0xFF, 0xFF), evaluator.Of(At(3)).ScaleColor);

        // Halfway between black and white, give or take a rounding step.
        ExcelColor middle = evaluator.Of(At(2)).ScaleColor;
        Assert.InRange((int)(middle.Argb & 0xFF), 0x7E, 0x81);
    }

    [Fact]
    public void ADataBarFillsInProportionToTheValue()
    {
        Worksheet sheet = Column(0, 25, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A3", new DataBar()));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0, evaluator.Of(At(1)).Bar!.Value.Fraction, 3);
        Assert.Equal(0.25, evaluator.Of(At(2)).Bar!.Value.Fraction, 3);
        Assert.Equal(1, evaluator.Of(At(3)).Bar!.Value.Fraction, 3);
    }

    [Fact]
    public void AnIconSetPutsAValueInTheBandItReaches()
    {
        Worksheet sheet = Column(0, 50, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A3", IconSet.TrafficLights));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0, evaluator.Of(At(1)).Icon!.Value.Index);
        Assert.Equal(1, evaluator.Of(At(2)).Icon!.Value.Index);
        Assert.Equal(2, evaluator.Of(At(3)).Icon!.Value.Index);
        Assert.Equal(3, evaluator.Of(At(3)).Icon!.Value.Count);
    }

    [Fact]
    public void AFormulaRuleWithoutAnEngineIsReportedRatherThanGuessedAt()
    {
        Worksheet sheet = Column(1);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression("A1", "$A1>0", DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.Contains("conditional-format-formula", evaluator.Unsupported);
    }

    [Fact]
    public void AFormulaRuleWithAnEnginePaints()
    {
        Worksheet sheet = Column(1, 0);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression("A1:A2", "$A1>0", DifferentialStyle.Good));

        // Stands in for the formula package: the rule is true where the cell holds something.
        ConditionalFormatEvaluator evaluator = Evaluator(
            sheet,
            (formula, cell) => CellValue.Boolean(sheet.Values.Get(cell.Row, cell.Column).ToSerialNumber() > 0));

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Empty(evaluator.Unsupported);
    }

    [Fact]
    public void ATimePeriodRuleComparesAgainstTheGivenToday()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = Today.ToDateTime(TimeOnly.MinValue);
        sheet["A2"] = Today.AddDays(-1).ToDateTime(TimeOnly.MinValue);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2")],
            Type = ConditionalFormatType.TimePeriod,
            TimePeriod = "today",
            Style = DifferentialStyle.Neutral,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
    }

    [Fact]
    public void ACellOutsideEveryRangeIsLeftAlone()
    {
        Worksheet sheet = Column(500, 500);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1", 1, DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.True(evaluator.Of(At(2)).IsEmpty);
    }

    [Fact]
    public void ASheetWithoutRulesHasNothingToApply()
    {
        Assert.False(Evaluator(Column(1)).HasRules);
    }
}
