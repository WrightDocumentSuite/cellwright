using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Writing input rules to the binary format and reading them back. The reader predates the
/// writer here, so a round trip checks the new half against the half the corpus already
/// exercises.
/// </summary>
public sealed class ValidationTests
{
    [Fact]
    public async Task ANumericRuleSurvivesAWriteAndRead()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Input");
        sheet.Validations.Add(DataValidation.WholeBetween("B2:B100", 1, 10));

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        DataValidation rule = Assert.Single(read.Validations);
        Assert.Equal(DataValidationType.Whole, rule.Type);
        Assert.Equal(DataValidationOperator.Between, rule.Operator);
        Assert.Equal("1", rule.Formula1);
        Assert.Equal("10", rule.Formula2);
        Assert.Equal("B2:B100", Assert.Single(rule.Ranges).ToString());
    }

    [Fact]
    public async Task AListOverARangeSurvives()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Input");
        Worksheet lists = original.Sheets.Add("Lists");
        lists["A1"] = CellValue.Text("Yes");
        sheet.Validations.Add(DataValidation.List("C2:C50", "Lists!$A$1:$A$5"));

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        DataValidation rule = Assert.Single(read.Validations);
        Assert.Equal(DataValidationType.List, rule.Type);
        Assert.Equal("Lists!$A$1:$A$5", rule.Formula1);
        Assert.True(rule.ShowDropDown);
    }

    /// <summary>
    /// A list of literal choices is not a formula in this format but a single string with the
    /// choices separated by a NUL, which is the one place the two representations differ.
    /// </summary>
    [Fact]
    public async Task AListOfLiteralChoicesSurvives()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Input");
        sheet.Validations.Add(DataValidation.List("A1:A20", "\"Low,Medium,High\""));

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        DataValidation rule = Assert.Single(read.Validations);
        Assert.Equal(DataValidationType.List, rule.Type);
        Assert.Equal("\"Low,Medium,High\"", rule.Formula1);
    }

    [Fact]
    public async Task MessagesAndTheErrorStyleSurvive()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Input");
        sheet.Validations.Add(new DataValidation
        {
            Ranges = DataValidation.ParseRanges("D1:D9"),
            Type = DataValidationType.Decimal,
            Operator = DataValidationOperator.GreaterThan,
            Formula1 = "0",
            AllowBlank = false,
            ErrorStyle = DataValidationErrorStyle.Warning,
            PromptTitle = "Amount",
            Prompt = "Enter a positive number",
            ErrorTitle = "Out of range",
            Error = "The amount has to be positive",
        });

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        DataValidation rule = Assert.Single(read.Validations);
        Assert.Equal(DataValidationOperator.GreaterThan, rule.Operator);
        Assert.False(rule.AllowBlank);
        Assert.Equal(DataValidationErrorStyle.Warning, rule.ErrorStyle);
        Assert.Equal("Amount", rule.PromptTitle);
        Assert.Equal("Enter a positive number", rule.Prompt);
        Assert.Equal("Out of range", rule.ErrorTitle);
        Assert.Equal("The amount has to be positive", rule.Error);
    }

    [Fact]
    public async Task ACustomRuleKeepsItsFormula()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Input");
        sheet.Validations.Add(DataValidation.Custom("A1:A10", "LEN(A1)<=5"));

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        DataValidation rule = Assert.Single(read.Validations);
        Assert.Equal(DataValidationType.Custom, rule.Type);
        Assert.Equal("LEN(A1)<=5", rule.Formula1);
    }

    private static async Task<Workbook> RoundTripAsync(Workbook workbook)
    {
        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        return await XlsWorkbook.LoadAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }
}
