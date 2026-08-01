using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Conditional formatting from both ends: rules built in the model have to come back after a
/// round trip and be understood by another library, and rules read from a file have to survive
/// even when the model has no way to describe them.
/// </summary>
public sealed class ConditionalFormatTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        for (int row = 1; row <= 10; row++)
        {
            sheet[row, 1] = row * 100;
            sheet[row, 2] = $"item-{row}";
        }

        return book;
    }

    [Fact]
    public async Task EveryRuleKindRoundTrips()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellBetween("A1:A10", 200, 400, DifferentialStyle.Neutral));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression("B1:B10", "$A1>800", DifferentialStyle.Good));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Containing("B1:B10", "item-3", DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Top("A1:A10", 3, DifferentialStyle.Good));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Duplicates("B1:B10", DifferentialStyle.Neutral));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A10", ColorScale.RedYellowGreen));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A10", new DataBar()));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A10", IconSet.TrafficLights));

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(9, loaded.ConditionalFormats.Count);
        Assert.Equal(
            [
                ConditionalFormatType.CellIs, ConditionalFormatType.CellIs, ConditionalFormatType.Expression,
                ConditionalFormatType.ContainsText, ConditionalFormatType.Top10, ConditionalFormatType.DuplicateValues,
                ConditionalFormatType.ColorScale, ConditionalFormatType.DataBar, ConditionalFormatType.IconSet,
            ],
            loaded.ConditionalFormats.Select(static rule => rule.Type));
        Assert.Equal(Enumerable.Range(1, 9), loaded.ConditionalFormats.Select(static rule => rule.Priority));
    }

    [Fact]
    public async Task ComparisonRuleKeepsItsOperatorBoundsAndRanges()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.CellBetween("A1:A10", 200, 400, DifferentialStyle.Bad));

        ConditionalFormatRule rule = Assert.Single((await SaveAndReloadAsync(book)).Sheets[0].ConditionalFormats);

        Assert.Equal(ConditionalOperator.Between, rule.Operator);
        Assert.Equal(["200", "400"], rule.Formulas);
        Assert.Equal(RangeRef.Parse("A1:A10"), Assert.Single(rule.Ranges));
    }

    [Fact]
    public async Task GradedRulesKeepTheirStopsAndColors()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A10", ColorScale.RedYellowGreen));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("B1:B10", new DataBar
        {
            Minimum = ConditionalValue.Number(0),
            Maximum = ConditionalValue.Number(1000),
            Color = ExcelColor.FromRgb(0x63, 0x8E, 0xC6),
            ShowValue = false,
        }));

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets[0];

        ColorScale scale = Assert.IsType<ColorScale>(loaded.ConditionalFormats[0].ColorScale);
        Assert.Equal(3, scale.Values.Count);
        Assert.Equal(ConditionalValueType.Minimum, scale.Values[0].Type);
        Assert.Equal(ConditionalValueType.Percentile, scale.Values[1].Type);
        Assert.Equal("50", scale.Values[1].Value);
        Assert.Equal(ExcelColor.FromRgb(0x63, 0xBE, 0x7B), scale.Colors[2]);

        DataBar bar = Assert.IsType<DataBar>(loaded.ConditionalFormats[1].DataBar);
        Assert.False(bar.ShowValue);
        Assert.Equal("1000", bar.Maximum.Value);
    }

    [Fact]
    public async Task ClosedXmlSeesTheRulesAndTheirFormatting()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        IXLConditionalFormat format = Assert.Single(interop.Worksheet("Data").ConditionalFormats);
        Assert.Equal(XLConditionalFormatType.CellIs, format.ConditionalFormatType);
        Assert.Equal(XLCFOperator.GreaterThan, format.Operator);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFFFC7CE)), format.Style.Fill.BackgroundColor);
    }

    [Fact]
    public async Task StopIfTrueAndPrioritySurvive()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A10")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.LessThan,
            Formulas = ["300"],
            Style = DifferentialStyle.Good,
            Priority = 5,
            StopIfTrue = true,
        });

        ConditionalFormatRule rule = Assert.Single((await SaveAndReloadAsync(book)).Sheets[0].ConditionalFormats);

        Assert.Equal(5, rule.Priority);
        Assert.True(rule.StopIfTrue);
    }

    [Fact]
    public async Task AddingARuleKeepsTheOnesTheModelCannotDescribe()
    {
        // An extension-only rule is what Excel writes for the modern icon sets. It has to come
        // back untouched even though the sheet's rules were regenerated around it.
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        sheet["A1"] = 1;
        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A10")],
            RawXml = """
                <cfRule type="dataBar" priority="9" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>
                <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"/></extLst></cfRule>
                """,
        });

        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B1:B10", 1, DifferentialStyle.Bad));
        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(2, loaded.ConditionalFormats.Count);
        Assert.True(loaded.ConditionalFormats[0].IsOpaque);
        Assert.Equal(ConditionalFormatType.CellIs, loaded.ConditionalFormats[1].Type);
    }

    [Fact]
    public async Task RulesReadFromAFileKeepTheirOriginalFormattingIndex()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));
        Workbook reloaded = await SaveAndReloadAsync(book);

        // Adding a second rule regenerates the block; the first must keep pointing at the
        // differential format that came with the file rather than losing its colors.
        reloaded.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellLessThan("A1:A10", 100, DifferentialStyle.Good));

        using var stream = new MemoryStream();
        await reloaded.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        List<IXLConditionalFormat> formats = [.. interop.Worksheet("Data").ConditionalFormats];
        Assert.Equal(2, formats.Count);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFFFC7CE)), formats[0].Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFC6EFCE)), formats[1].Style.Fill.BackgroundColor);
    }
}
