using System.Globalization;
using Xunit;

namespace Cellwright.Tests.Markdown;

/// <summary>
/// Turning a workbook into GitHub Flavored Markdown. There is no Markdown parser here to act as
/// an oracle, so the tables are checked against exact text, and the one structural rule of the
/// table extension — the header and the delimiter hold the same number of cells — is checked
/// on its own.
/// </summary>
public sealed class MarkdownExportTests
{
    /// <summary>
    /// The whole conversion at once, and the example <c>docs/markdown.md</c> shows: a heading per
    /// sheet, the used range as a table, values through their number formats, and the alignment
    /// of the header carried into the delimiter row.
    /// </summary>
    [Fact]
    public void ASheetBecomesASectionWithATable()
    {
        var book = Workbook.Create();
        book.Properties.Title = "Отчёт за квартал";
        Worksheet sheet = book.Sheets.Add("Продажи");
        sheet["A1"] = "Товар";
        sheet["B1"] = "Сумма";
        sheet["C1"] = "Дата";
        sheet.Cell("B1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Right));

        CellStyle money = CellStyle.Default.NumberFormat("#,##0.00");
        CellStyle day = CellStyle.Default.NumberFormat("yyyy-mm-dd");
        sheet["A2"] = "Кофе";
        sheet.Cell("B2").SetValue(1234.5).SetStyle(money);
        sheet.Cell("C2").SetValue(new DateTime(2026, 7, 25)).SetStyle(day);
        sheet["A3"] = "Чай";
        sheet.Cell("B3").SetValue(990).SetStyle(money);
        sheet.Cell("C3").SetValue(new DateTime(2026, 7, 26)).SetStyle(day);

        string markdown = book.ToMarkdown().Text;

        Assert.Equal(
            Lf("""
            # Отчёт за квартал

            ## Продажи

            | Товар | Сумма | Дата |
            | --- | ---: | --- |
            | Кофе | 1,234.50 | 2026-07-25 |
            | Чай | 990.00 | 2026-07-26 |

            """),
            markdown);
    }

    /// <summary>
    /// Without a title in the document properties the heading falls back to a fixed word, and an
    /// option overrides both — the file name a caller wants on top is rarely the one Excel kept.
    /// </summary>
    [Fact]
    public void TheHeadingComesFromTheTitleOrTheOptions()
    {
        var book = Workbook.Create();
        book.Sheets.Add("One");

        Assert.StartsWith("# Workbook\n", book.ToMarkdown().Text, StringComparison.Ordinal);
        Assert.StartsWith(
            "# Свод\n",
            book.ToMarkdown(new MarkdownExportOptions { WorkbookTitle = "Свод" }).Text,
            StringComparison.Ordinal);
    }

    [Fact]
    public void HeadingsPreserveMarkdownPunctuationLiterally()
    {
        var book = Workbook.Create();
        book.Properties.Title = "*Q1* &copy; <report> #";
        book.Sheets.Add("Sales &copy; <Q1> #")["A1"] = "x";

        string markdown = book.ToMarkdown().Text;

        Assert.StartsWith(@"# \*Q1\* \&copy\; \<report\> \#", markdown, StringComparison.Ordinal);
        Assert.Contains(@"## Sales \&copy\; \<Q1\> \#", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void AHeadingCrLfBecomesOneSpace()
    {
        var book = Workbook.Create();
        book.Sheets.Add("One")["A1"] = "x";

        string markdown = book.ToMarkdown(new MarkdownExportOptions { WorkbookTitle = "First\r\nSecond" }).Text;

        Assert.StartsWith("# First Second\n", markdown, StringComparison.Ordinal);
    }

    /// <summary>
    /// Column alignment is the delimiter row's only job beyond separating the header, and it is
    /// read off the cells of the first row: that is the header, and it is what the eye follows.
    /// </summary>
    [Fact]
    public void TheHeaderAlignmentChoosesTheColumnAlignment()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Aligned");
        sheet["A1"] = "left";
        sheet["B1"] = "center";
        sheet["C1"] = "right";
        sheet["D1"] = "general";
        sheet.Cell("A1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Left));
        sheet.Cell("B1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Center));
        sheet.Cell("C1").SetStyle(CellStyle.Default.Align(HorizontalAlignment.Right));

        Assert.Contains("| :--- | :---: | ---: | --- |", book.ToMarkdown().Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A pipe ends a cell and a backslash swallows what follows it, so both have to be escaped;
    /// a table cell cannot hold a line break, so it becomes one. Everything else goes through as
    /// it is, which is why the emphasis in the last cell survives as emphasis.
    /// </summary>
    [Fact]
    public void WhatWouldBreakTheTableIsEscaped()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Escapes");
        sheet["A1"] = "a|b";
        sheet["B1"] = @"C:\temp";
        sheet["C1"] = "first\r\nsecond";
        sheet["D1"] = "*note*";

        string markdown = book.ToMarkdown().Text;

        Assert.Contains(
            @"| a\|b | C:\\temp | first<br>second | *note* |", markdown, StringComparison.Ordinal);
        Assert.Equal(4, CountCells(TableRows(markdown)[0]));
    }

    [Fact]
    public void LeadingAndTrailingCellWhitespaceSurvivesGfmTrimming()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Spaces");
        sheet["A1"] = " left ";
        sheet["B1"] = "\ttab\t";

        Assert.Contains(
            "| &nbsp;left&nbsp; | &#9;tab&#9; |",
            book.ToMarkdown().Text,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// The one rule the extension enforces: a header and a delimiter of different widths is not
    /// a table at all. Escaped pipes must not count towards either.
    /// </summary>
    [Fact]
    public void EveryTableHasAsManyDelimiterCellsAsHeaderCells()
    {
        var book = Workbook.Create();
        Worksheet wide = book.Sheets.Add("Wide");
        for (int column = 1; column <= 7; column++)
            wide.Cell(1, column).SetValue($"c{column}|{column}");
        wide["A2"] = "value";

        Worksheet narrow = book.Sheets.Add("Narrow");
        narrow["A1"] = "only";

        string markdown = book.ToMarkdown().Text;
        List<string> rows = TableRows(markdown);

        Assert.Equal(7, CountCells(rows[0]));
        Assert.Equal(7, CountCells(rows[1]));
        Assert.Equal(1, CountCells(rows[rows.Count - 2]));
        Assert.Equal(1, CountCells(rows[rows.Count - 1]));
    }

    /// <summary>
    /// The point of the export: a cell reads the way Excel shows it, not the way its value is
    /// stored. A date is a number in the file, and a currency is the same number as a percentage.
    /// </summary>
    [Fact]
    public void ValuesAreRenderedThroughTheirNumberFormat()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Formats");
        sheet["A1"] = "date";
        sheet["B1"] = "money";
        sheet["C1"] = "share";
        sheet.Cell("A2")
            .SetValue(new DateTime(2026, 7, 25))
            .SetStyle(CellStyle.Default.NumberFormat("yyyy-mm-dd"));
        sheet.Cell("B2").SetValue(1234.5).SetStyle(CellStyle.Default.NumberFormat("$#,##0.00"));
        sheet.Cell("C2").SetValue(0.256).SetStyle(CellStyle.Default.NumberFormat("0.0%"));

        Assert.Contains("| 2026-07-25 | $1,234.50 | 25.6% |", book.ToMarkdown().Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// The library renders invariantly by default so that a document reads the same wherever it
    /// was produced; a caller who wants the local shape of a number asks for it.
    /// </summary>
    [Fact]
    public void TheCultureIsInvariantUnlessOneIsAskedFor()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Culture");
        sheet["A1"] = "money";
        sheet.Cell("A2").SetValue(1234.5).SetStyle(CellStyle.Default.NumberFormat("#,##0.00"));

        Assert.Contains("| 1,234.50 |", book.ToMarkdown().Text, StringComparison.Ordinal);
        Assert.Contains(
            "| 1.234,50 |",
            book.ToMarkdown(new MarkdownExportOptions { Culture = CultureInfo.GetCultureInfo("de-DE") }).Text,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// A formula contributes the value Excel last computed for it. The text of the formula is
    /// what the sheet is made of, not what it shows.
    /// </summary>
    [Fact]
    public void AFormulaContributesItsValueRatherThanItsText()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Formulas");
        sheet["A1"] = "total";
        Cell total = sheet.Cell("A2");
        total.SetFormula("SUM(B1:B9)");
        total.SetComputedValue(CellValue.Number(42));

        string markdown = book.ToMarkdown().Text;

        Assert.Contains("| 42 |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("SUM", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void AnExternalLinkKeepsItsLocationAsAFragment()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Links");
        sheet["A1"] = "Section";
        sheet.Hyperlinks.Add(Hyperlink.Load(
            RangeRef.Parse("A1"), "https://example.com/book.xlsx", "Sheet2!A1", tooltip: null, display: null));

        Assert.Contains(
            "[Section](https://example.com/book.xlsx#Sheet2!A1)",
            book.ToMarkdown().Text,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("JaVaScRiPt&colon;alert(1)")]
    [InlineData("java&#10;script:alert(1)")]
    [InlineData("data:text/html,unsafe")]
    [InlineData("vbscript:msgbox(1)")]
    [InlineData("https://example.com/first\nsecond")]
    public void AnUnsafeHyperlinkIsRenderedAsPlainText(string destination)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Links");
        sheet["A1"] = "Untrusted";
        sheet.Hyperlinks.Add(RangeRef.Parse("A1"), destination);

        string markdown = book.ToMarkdown().Text;

        Assert.Contains("| Untrusted |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[Untrusted]", markdown, StringComparison.Ordinal);
    }

    /// <summary>
    /// A link keeps the text the cell shows. A destination holding a parenthesis goes into the
    /// pointy-bracket form CommonMark provides for one, or the link would end at the parenthesis.
    /// </summary>
    [Fact]
    public void AHyperlinkBecomesALink()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Links");
        sheet["A1"] = "site";
        sheet["B1"] = "wiki";
        sheet["A2"] = "Пример";
        sheet["B2"] = "Статья";
        sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com/a");
        sheet.Hyperlinks.Add(RangeRef.Parse("B2"), "https://example.com/Article_(disambiguation)");

        Assert.Contains(
            "| [Пример](https://example.com/a) | [Статья](<https://example.com/Article_(disambiguation)>) |",
            book.ToMarkdown().Text,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// A link over a blank cell has nothing to show but what the file recorded beside it, and
    /// failing that the destination — an invisible link would be worse than a verbose one.
    /// </summary>
    [Fact]
    public void ALinkOverABlankCellFallsBackToTheDestination()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Blank");
        sheet["A1"] = "link";
        sheet.Cell("A2").SetStyle(CellStyle.Default.Bold());
        sheet.Hyperlinks.Add(RangeRef.Parse("A2"), "https://example.com/");

        Assert.Contains(
            "| [https://example.com/](https://example.com/) |", book.ToMarkdown().Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// Markdown has no spanning cell, so a merge puts its value in the corner it belongs to and
    /// leaves the rest of the range empty — which is what the sheet looks like anyway.
    /// </summary>
    [Fact]
    public void AMergeFillsOnlyItsTopLeftCell()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Merged");
        sheet["A1"] = "left";
        sheet["B1"] = "right";
        sheet["A2"] = "Итого";
        sheet["B2"] = "ignored";
        sheet.Merge(RangeRef.Parse("A2:B2"));

        string markdown = book.ToMarkdown().Text;

        Assert.Contains("| Итого |  |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("ignored", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptySheetKeepsItsHeading()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Пусто");

        Assert.Equal(Lf("""
            # Workbook

            ## Пусто

            (пусто)

            """),
            book.ToMarkdown().Text);
    }

    /// <summary>Hidden sheets are out of sight in Excel, and out of the document here.</summary>
    [Fact]
    public void HiddenSheetsAreLeftOut()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Shown")["A1"] = "visible";
        Worksheet hidden = book.Sheets.Add("Hidden");
        hidden["A1"] = "secret";
        hidden.Visibility = SheetVisibility.VeryHidden;

        string markdown = book.ToMarkdown().Text;

        Assert.Contains("## Shown", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", markdown, StringComparison.Ordinal);
    }

    /// <summary>
    /// A sheet whose first row is data rather than a header still needs one, and column letters
    /// are the header Excel itself shows.
    /// </summary>
    [Fact]
    public void ColumnLettersKeepEveryRowAsData()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        sheet["A1"] = "Кофе";
        sheet["B1"] = 1;

        string markdown = book
            .ToMarkdown(new MarkdownExportOptions { HeaderRow = MarkdownHeaderRow.ColumnLetters })
            .Text;

        Assert.Contains(Lf("""
            | A | B |
            | --- | --- |
            | Кофе | 1 |
            """),
            markdown);
    }

    private static string Lf(string text) => text.ReplaceLineEndings("\n");

    /// <summary>Every line of the document that is part of a table.</summary>
    private static List<string> TableRows(string markdown) =>
        [.. markdown.Split('\n').Where(static line => line.StartsWith('|'))];

    /// <summary>
    /// Cells of a table row, counted the way a reader of the format counts them: an escaped pipe
    /// is content, and the pipes bounding the row are not separators.
    /// </summary>
    private static int CountCells(string row)
    {
        int pipes = 0;
        for (int i = 0; i < row.Length; i++)
        {
            if (row[i] == '\\')
                i++;
            else if (row[i] == '|')
                pipes++;
        }

        return pipes - 1;
    }
}
