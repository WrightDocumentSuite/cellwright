using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reading conditional formatting out of the binary format. The fixtures are the POI test
/// workbooks whose rules that project's own suite states cell by cell, so the expectations
/// here are not this library's own guesses.
/// </summary>
public sealed class ConditionalFormatTests
{
    private static string? Find(string name) =>
        XlsCorpus.Files.FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    private static async Task<Worksheet?> LoadSheetAsync(string file, string sheetName)
    {
        string? path = Find(file);
        if (path is null)
            return null;

        Workbook workbook = await XlsWorkbook.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
        return workbook.Sheets.Find(sheetName);
    }

    [Fact]
    public async Task ComparisonRulesKeepTheirOperatorAndBounds()
    {
        Worksheet? sheet = await LoadSheetAsync("WithConditionalFormatting.xls", "CF");
        Assert.SkipWhen(sheet is null, "WithConditionalFormatting.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        // A1:A8 carries two rules: greater than 3, and less than -3.
        List<ConditionalFormatRule> overA = [.. sheet!.ConditionalFormats.Where(
            static rule => rule.Ranges.Any(range => range.ToString() == "A1:A8"))];
        Assert.Equal(2, overA.Count);

        ConditionalFormatRule greater = overA[0];
        Assert.Equal(ConditionalFormatType.CellIs, greater.Type);
        Assert.Equal(ConditionalOperator.GreaterThan, greater.Operator);
        Assert.Equal("3", Assert.Single(greater.Formulas));
        Assert.True(greater.Style?.Bold);

        ConditionalFormatRule less = overA[1];
        Assert.Equal(ConditionalOperator.LessThan, less.Operator);
        Assert.Equal("-3", Assert.Single(less.Formulas));
        Assert.True(less.Style?.Bold);
        Assert.True(less.Style?.Italic);
    }

    [Fact]
    public async Task AFormulaRuleKeepsItsExpression()
    {
        Worksheet? sheet = await LoadSheetAsync("WithConditionalFormatting.xls", "CF");
        Assert.SkipWhen(sheet is null, "WithConditionalFormatting.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        ConditionalFormatRule rule = Assert.Single(
            sheet!.ConditionalFormats, static r => r.Type == ConditionalFormatType.Expression);
        Assert.Equal("B9", Assert.Single(rule.Ranges).ToString());
        Assert.Equal("$A$8>5", Assert.Single(rule.Formulas));
    }

    /// <summary>
    /// The graded rule types arrived with Excel 2007 and are stored in records of their own.
    /// The fixture has one of each, over single columns.
    /// </summary>
    [Fact]
    public async Task GradedRulesAreRead()
    {
        string? path = Find("NewStyleConditionalFormattings.xls");
        Assert.SkipWhen(path is null, "NewStyleConditionalFormattings.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);
        Worksheet sheet = workbook.Sheets[0];

        // The fixture has two three-stop scales; the red-yellow-green one is Excel's own.
        ConditionalFormatRule scale = sheet.ConditionalFormats.First(static r =>
            r.Type == ConditionalFormatType.ColorScale && r.ColorScale!.Colors[0].Argb == 0xFFF8696Bu);
        Assert.Equal(ConditionalValueType.Minimum, scale.ColorScale!.Values[0].Type);
        Assert.Equal(ConditionalValueType.Percentile, scale.ColorScale.Values[1].Type);
        Assert.Equal(ConditionalValueType.Maximum, scale.ColorScale.Values[2].Type);
        Assert.Equal(3, scale.ColorScale.Colors.Count);

        // Excel's own red-yellow-green scale, which is what the fixture uses.
        Assert.Equal(0xFFF8696Bu, scale.ColorScale.Colors[0].Argb);
        Assert.Equal(0xFF63BE7Bu, scale.ColorScale.Colors[2].Argb);

        ConditionalFormatRule bar = sheet.ConditionalFormats.First(static r => r.Type == ConditionalFormatType.DataBar);
        Assert.Equal(ConditionalValueType.Minimum, bar.DataBar!.Minimum.Type);
        Assert.Equal(ConditionalValueType.Maximum, bar.DataBar.Maximum.Type);
        Assert.Equal(0xFF63C384u, bar.DataBar.Color.Argb);

        ConditionalFormatRule icons = sheet.ConditionalFormats.First(
            static r => r.Type == ConditionalFormatType.IconSet && r.IconSet!.Name == "3TrafficLights1");
        Assert.Equal(3, icons.IconSet!.Values.Count);
        Assert.Equal("0", icons.IconSet.Values[0].Value);
        Assert.Equal("33", icons.IconSet.Values[1].Value);
        Assert.Equal("67", icons.IconSet.Values[2].Value);
    }

    /// <summary>
    /// Rules read from a file belong to the file until something changes them, so converting a
    /// workbook to the modern format has to carry them over.
    /// </summary>
    [Fact]
    public async Task RulesSurviveConversionToTheModernFormat()
    {
        string? path = Find("WithConditionalFormatting.xls");
        Assert.SkipWhen(path is null, "WithConditionalFormatting.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook legacy = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);

        using var buffer = new MemoryStream();
        await legacy.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        Workbook modern = await Workbook.LoadAsync(
            buffer, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Worksheet sheet = modern.Sheets.Find("CF")!;
        Assert.Contains(sheet.ConditionalFormats, static rule =>
            rule.Type == ConditionalFormatType.Expression && rule.Formulas.Contains("$A$8>5"));
    }
}
