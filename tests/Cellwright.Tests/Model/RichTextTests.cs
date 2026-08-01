using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Character formatting inside a cell: reading the runs a file carries, writing new ones, and
/// what happens to them when the text they describe is replaced.
/// </summary>
public sealed class RichTextTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        // The runs are written as rPr, whose schema differs from a style's font in one element.
        // The official validator is what says the difference was honoured.
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

    [Fact]
    public void TheVisibleTextIsEveryRunEndToEnd()
    {
        RichText text = RichText.Empty
            .Append("Total: ")
            .Append("42", FontStyle.Default with { Bold = true });

        Assert.Equal("Total: 42", text.Text);
        Assert.Equal(2, text.Count);
        Assert.True(text.IsFormatted);
    }

    [Fact]
    public void TextWhoseRunsAllFollowTheCellIsNotFormatted()
    {
        RichText text = RichText.Empty.Append("one ").Append("two");

        Assert.False(text.IsFormatted);
        Assert.Equal("one two", text.Text);
    }

    [Fact]
    public async Task RunsSurviveASaveAndReload()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty
            .Append("plain ")
            .Append("bold", FontStyle.Default with { Bold = true }));

        Workbook reloaded = await SaveAndReloadAsync(book);
        RichText? runs = reloaded.Sheets[0].Cell("A1").RichText;

        Assert.Equal("plain bold", reloaded.Sheets[0]["A1"].GetText());
        Assert.NotNull(runs);
        Assert.Equal(2, runs.Count);
        Assert.Equal("plain ", runs[0].Text);
        Assert.Null(runs[0].Font);
        Assert.Equal("bold", runs[1].Text);
        Assert.True(runs[1].Font!.Bold);
    }

    [Fact]
    public async Task EveryFontPropertyOfARunSurvives()
    {
        var font = FontStyle.Default with
        {
            Name = "Georgia",
            Size = 14,
            Bold = true,
            Italic = true,
            Strikethrough = true,
            Underline = UnderlineStyle.Double,
            VerticalAlign = FontVerticalAlign.Superscript,
            Color = ExcelColor.FromRgb(0x9C, 0x00, 0x06),
        };

        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty.Append("styled", font));

        Workbook reloaded = await SaveAndReloadAsync(book);
        FontStyle? read = reloaded.Sheets[0].Cell("A1").RichText![0].Font;

        Assert.NotNull(read);
        Assert.Equal("Georgia", read.Name);
        Assert.Equal(14, read.Size);
        Assert.True(read.Bold);
        Assert.True(read.Italic);
        Assert.True(read.Strikethrough);
        Assert.Equal(UnderlineStyle.Double, read.Underline);
        Assert.Equal(FontVerticalAlign.Superscript, read.VerticalAlign);
        Assert.Equal(ExcelColor.FromRgb(0x9C, 0x00, 0x06), read.Color);
    }

    [Fact]
    public async Task ARunKeepsTheSpacesAtItsEdges()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty
            .Append("  leading", FontStyle.Default with { Bold = true })
            .Append("trailing  ", FontStyle.Default with { Italic = true }));

        Workbook reloaded = await SaveAndReloadAsync(book);
        RichText runs = reloaded.Sheets[0].Cell("A1").RichText!;

        Assert.Equal("  leading", runs[0].Text);
        Assert.Equal("trailing  ", runs[1].Text);
    }

    [Fact]
    public async Task TextWithNoFormattingIsStoredAsAnOrdinaryString()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty.Append("one ").Append("two"));

        Workbook reloaded = await SaveAndReloadAsync(book);

        Assert.Equal("one two", reloaded.Sheets[0]["A1"].GetText());
        Assert.Null(reloaded.Sheets[0].Cell("A1").RichText);
    }

    [Fact]
    public void ReplacingTheValueDropsTheFormatting()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Cell cell = sheet.Cell("A1");
        cell.SetRichText(RichText.Empty.Append("bold", FontStyle.Default with { Bold = true }));
        Assert.NotNull(cell.RichText);

        cell.SetValue(CellValue.Text("something else"));

        Assert.Null(cell.RichText);
        Assert.Equal("something else", cell.Value.GetText());
    }

    [Fact]
    public void ClearingTheCellDropsTheFormatting()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Cell cell = sheet.Cell("A1");
        cell.SetRichText(RichText.Empty.Append("bold", FontStyle.Default with { Bold = true }));

        cell.Clear();

        Assert.Null(cell.RichText);
        Assert.True(cell.Value.IsBlank);
    }

    [Fact]
    public void ACellWithoutRunsHasNoRichText()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "plain";

        Assert.Null(sheet.Cell("A1").RichText);
    }

    [Fact]
    public async Task SameTextWithDifferentRunsStaysDistinct()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty.Append("same", FontStyle.Default with { Bold = true }));
        sheet.Cell("B1").SetRichText(RichText.Empty.Append("same", FontStyle.Default with { Italic = true }));

        Workbook reloaded = await SaveAndReloadAsync(book);
        Worksheet read = reloaded.Sheets[0];

        Assert.True(read.Cell("A1").RichText![0].Font!.Bold);
        Assert.False(read.Cell("A1").RichText![0].Font!.Italic);
        Assert.True(read.Cell("B1").RichText![0].Font!.Italic);
        Assert.False(read.Cell("B1").RichText![0].Font!.Bold);
    }

    [Fact]
    public async Task SettingRichTextOverAFormulaRemovesIt()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Cell cell = sheet.Cell("A1");
        cell.SetFormula("1+1");

        cell.SetRichText(RichText.Empty.Append("text", FontStyle.Default with { Bold = true }));

        Assert.Null(cell.Formula);

        Workbook reloaded = await SaveAndReloadAsync(book);
        Assert.Null(reloaded.Sheets[0].Cell("A1").Formula);
        Assert.Equal("text", reloaded.Sheets[0]["A1"].GetText());
    }

    [Fact]
    public void EmptyRichTextWritesAnEmptyString()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Cell("A1").SetRichText(RichText.Empty);

        Assert.Equal(string.Empty, sheet.Cell("A1").Value.GetText());
        Assert.Null(sheet.Cell("A1").RichText);
    }

    [Fact]
    public void APlainStringBecomesOneRun()
    {
        RichText text = RichText.FromPlain("just text");

        Assert.Single(text);
        Assert.Equal("just text", text[0].Text);
        Assert.Null(text[0].Font);
        Assert.False(text.IsFormatted);
    }
}
