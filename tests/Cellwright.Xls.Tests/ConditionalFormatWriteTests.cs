using Cellwright.Styles;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Writing conditional formatting to .xls and reading it back: the rules the classic records can
/// carry, the ones that are lowered to a formula on the way, and the ones that are reported
/// rather than approximated.
/// </summary>
public sealed class ConditionalFormatWriteTests
{
    private static async Task<Worksheet> RoundTripAsync(Workbook book, List<WorkbookWarning>? warnings = null)
    {
        using var stream = new MemoryStream();
        await XlsWorkbook.SaveAsync(
            book,
            stream,
            new XlsSaveOptions { OnWarning = warnings is null ? null : warnings.Add },
            TestContext.Current.CancellationToken);

        stream.Position = 0;
        Workbook read = await XlsWorkbook.LoadAsync(stream, options: null, TestContext.Current.CancellationToken);

        return read.Sheets[0];
    }

    private static Worksheet Build(Workbook book)
    {
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int row = 1; row <= 5; row++)
            sheet[row, 1] = row * 100;

        return sheet;
    }

    [Fact]
    public async Task AComparisonRuleSurvivesTheRoundTrip()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, new DifferentialStyle { Bold = true }));

        Worksheet read = await RoundTripAsync(book);

        ConditionalFormatRule rule = Assert.Single(read.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.CellIs, rule.Type);
        Assert.Equal(ConditionalOperator.GreaterThan, rule.Operator);
        Assert.Equal("A1:A5", rule.Ranges[0].ToString());
        Assert.True(rule.Style!.Bold);
    }

    [Fact]
    public async Task ABetweenRuleKeepsBothOfItsOperands()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellBetween(
            "A1:A5", 150, 350, DifferentialStyle.Good));

        Worksheet read = await RoundTripAsync(book);

        ConditionalFormatRule rule = Assert.Single(read.ConditionalFormats);
        Assert.Equal(ConditionalOperator.Between, rule.Operator);
        Assert.Equal(2, rule.Formulas.Count);
        Assert.Equal("150", rule.Formulas[0]);
        Assert.Equal("350", rule.Formulas[1]);
    }

    [Fact]
    public async Task AFormulaRuleSurvivesAsItself()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression(
            "A1:A5", "A1>250", DifferentialStyle.Bad));

        Worksheet read = await RoundTripAsync(book);

        ConditionalFormatRule rule = Assert.Single(read.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.Expression, rule.Type);
        Assert.Equal("A1>250", rule.Formulas[0]);
    }

    [Fact]
    public async Task TheColoursOfAFormatSurvive()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, DifferentialStyle.Bad));

        Worksheet read = await RoundTripAsync(book);
        DifferentialStyle style = read.ConditionalFormats[0].Style!;

        Assert.Equal(DifferentialStyle.Bad.FontColor.Argb & 0xFFFFFF, style.FontColor.Argb & 0xFFFFFF);
        Assert.Equal(DifferentialStyle.Bad.FillColor.Argb & 0xFFFFFF, style.FillColor.Argb & 0xFFFFFF);
    }

    [Fact]
    public async Task APropertyTheRuleNeverMentionedIsNotResetOnTheWayBack()
    {
        // The "unchanged" bits are the whole point of a differential format: a rule that only
        // sets the fill must not come back also claiming to unset bold.
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0xC7, 0xCE) }));

        Worksheet read = await RoundTripAsync(book);
        DifferentialStyle style = read.ConditionalFormats[0].Style!;

        Assert.Null(style.Bold);
        Assert.Null(style.Italic);
        Assert.Null(style.Underline);
        Assert.Null(style.Strikethrough);
        Assert.True(style.FontColor.IsAuto);
        Assert.False(style.FillColor.IsAuto);
    }

    [Fact]
    public async Task AnItalicRuleComesBackItalicAndNotBold()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, new DifferentialStyle { Italic = true, Strikethrough = true }));

        Worksheet read = await RoundTripAsync(book);
        DifferentialStyle style = read.ConditionalFormats[0].Style!;

        Assert.True(style.Italic);
        Assert.True(style.Strikethrough);
        Assert.Null(style.Bold);
    }

    [Fact]
    public async Task ANumberFormatOfARuleSurvives()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, new DifferentialStyle { NumberFormat = "0.00%" }));

        Worksheet read = await RoundTripAsync(book);

        Assert.Equal("0.00%", read.ConditionalFormats[0].Style!.NumberFormat);
    }

    [Fact]
    public async Task ATextRuleIsLoweredToTheFormulaItMeans()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Containing(
            "A1:A5", "overdue", DifferentialStyle.Bad));

        var warnings = new List<WorkbookWarning>();
        Worksheet read = await RoundTripAsync(book, warnings);

        // The classic records have no "contains text" condition, so it arrives as the formula
        // Excel itself writes for that rule — which is not a loss, just a different spelling.
        ConditionalFormatRule rule = Assert.Single(read.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.Expression, rule.Type);
        Assert.Contains("SEARCH", rule.Formulas[0], StringComparison.Ordinal);
        Assert.Contains("overdue", rule.Formulas[0], StringComparison.Ordinal);
        Assert.DoesNotContain(warnings, w => w.Code == "ConditionalFormatsNotWritten");
    }

    [Fact]
    public async Task ABlanksRuleIsLoweredToo()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A5")],
            Type = ConditionalFormatType.ContainsBlanks,
            Style = DifferentialStyle.Neutral,
        });

        Worksheet read = await RoundTripAsync(book);

        Assert.Contains("TRIM", read.ConditionalFormats[0].Formulas[0], StringComparison.Ordinal);
    }

    [Fact]
    public async Task AGradedRuleIsReportedRatherThanApproximated()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A5", ColorScale.RedYellowGreen));

        var warnings = new List<WorkbookWarning>();
        Worksheet read = await RoundTripAsync(book, warnings);

        Assert.Empty(read.ConditionalFormats);
        Assert.Contains(warnings, w => w.Code == "ConditionalFormatsNotWritten");
    }

    [Fact]
    public async Task TheRulesThatCanBeWrittenAreWrittenEvenBesideOnesThatCannot()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A5", new DataBar()));

        var warnings = new List<WorkbookWarning>();
        Worksheet read = await RoundTripAsync(book, warnings);

        Assert.Single(read.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.CellIs, read.ConditionalFormats[0].Type);
        Assert.Contains(warnings, w => w.Code == "ConditionalFormatsNotWritten");
    }

    [Fact]
    public async Task MoreThanThreeRulesOverOneRangeAreSplitAcrossGroups()
    {
        // A CondFmt counts its rules in a field that allows three, so a fourth needs a group of
        // its own rather than to be dropped.
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        for (int i = 1; i <= 5; i++)
        {
            sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
                "A1:A5", i * 100, new DifferentialStyle { Bold = true }));
        }

        Worksheet read = await RoundTripAsync(book);

        Assert.Equal(5, read.ConditionalFormats.Count);
    }

    [Fact]
    public async Task RulesOverDifferentRangesBecomeDifferentGroups()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Build(book);
        sheet[1, 2] = 1;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A5", 250, DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellLessThan(
            "B1:B5", 5, DifferentialStyle.Good));

        Worksheet read = await RoundTripAsync(book);

        Assert.Equal(2, read.ConditionalFormats.Count);
        Assert.Equal("A1:A5", read.ConditionalFormats[0].Ranges[0].ToString());
        Assert.Equal("B1:B5", read.ConditionalFormats[1].Ranges[0].ToString());
    }

    [Fact]
    public async Task ASheetWithoutRulesSaysNothingAboutThem()
    {
        Workbook book = Workbook.Create();
        Build(book);

        var warnings = new List<WorkbookWarning>();
        Worksheet read = await RoundTripAsync(book, warnings);

        Assert.Empty(read.ConditionalFormats);
        Assert.DoesNotContain(warnings, w => w.Code == "ConditionalFormatsNotWritten");
    }
}
