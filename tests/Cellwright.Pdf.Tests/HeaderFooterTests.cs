using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Headers and footers: the codes they are written in, which page gets which, and where on the
/// paper they land.
/// </summary>
public sealed class HeaderFooterTests
{
    private static Workbook Book(Action<PageSetup> setup, int rows = 1)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");

        for (int row = 1; row <= rows; row++)
            sheet.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"§{row}"));

        setup(sheet.PageSetup);
        return book;
    }

    [Fact]
    public void APlainHeaderIsPrintedAtTheTop()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&CQuarterly report" });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Quarterly report", rendered.Text(0), StringComparison.Ordinal);

        (double X, double Y)? header = rendered.Find("Q");
        (double X, double Y)? body = rendered.Find("§");

        Assert.NotNull(header);
        Assert.NotNull(body);
        Assert.True(header!.Value.Y > body!.Value.Y, "the header was not above the body");
    }

    [Fact]
    public void AFooterIsPrintedAtTheBottom()
    {
        Workbook book = Book(setup => setup.Footer = new HeaderFooterText { Default = "&LThe end" });

        using Rendered rendered = Rendered.Of(book);

        (double X, double Y)? footer = rendered.Find("T");
        (double X, double Y)? body = rendered.Find("§");

        Assert.NotNull(footer);
        Assert.NotNull(body);
        Assert.True(footer!.Value.Y < body!.Value.Y - 500, "the footer was not at the foot of the page");
    }

    [Fact]
    public void TheThreeSectionsGoLeftMiddleAndRight()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&LLeft&CMid&RRight" });

        using Rendered rendered = Rendered.Of(book);

        (double X, double Y)? left = rendered.Find("L");
        (double X, double Y)? middle = rendered.Find("M");
        (double X, double Y)? right = rendered.Find("R");

        Assert.NotNull(left);
        Assert.NotNull(middle);
        Assert.NotNull(right);
        Assert.True(left!.Value.X < middle!.Value.X, "the middle section was not right of the left one");
        Assert.True(middle.Value.X < right!.Value.X, "the right section was not right of the middle one");
    }

    [Fact]
    public void ThePageNumberAndCountAreFilledIn()
    {
        Workbook book = Book(
            setup => setup.Footer = new HeaderFooterText { Default = "&CPage &P of &N" },
            rows: 120);

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.PageCount >= 2);
        Assert.Contains($"Page 1 of {rendered.PageCount}", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains($"Page 2 of {rendered.PageCount}", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void TheSheetNameIsFilledIn()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&L&A" });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Report", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void AnOffsetPageNumberCountsFromWhereItSays()
    {
        Workbook book = Book(setup => setup.Footer = new HeaderFooterText { Default = "&C&P+10" });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("11", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void ADoubledAmpersandIsOneAmpersand()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&LSales && costs" });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Sales & costs", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void TheFirstPageTakesItsOwnHeaderWhenTheSheetAsks()
    {
        Workbook book = Book(
            setup =>
            {
                setup.DifferentFirstPage = true;
                setup.Header = new HeaderFooterText { Default = "&CEvery page", First = "&CTitle page" };
            },
            rows: 120);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Title page", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("Every page", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void EvenPagesTakeTheirOwnHeaderWhenTheSheetAsks()
    {
        Workbook book = Book(
            setup =>
            {
                setup.DifferentOddEven = true;
                setup.Header = new HeaderFooterText { Default = "&COdd side", Even = "&CEven side" };
            },
            rows: 120);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Odd side", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("Even side", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void AFontCodeChangesTheFaceOfWhatFollows()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&L&\"Courier New\"&14Big" });

        using Rendered rendered = Rendered.Of(book);
        Inkwright.Text.PdfLetter big = rendered.Letters().First(letter => letter.Text == "B");

        Assert.True(big.Width > 6, $"the size code was ignored: the letter is {big.Width:0.0} wide");
    }

    [Fact]
    public void AFontCodeLeavesWhatCameBeforeItAlone()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&Lsmall&20BIG" });

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<Inkwright.Text.PdfLetter> letters = rendered.Letters();

        double small = letters.First(letter => letter.Text == "s").Width;
        double big = letters.First(letter => letter.Text == "B").Width;

        Assert.True(big > small * 1.5, $"the size code reached backwards: 's' is {small:0.0}, 'B' is {big:0.0}");
    }

    [Fact]
    public void AnEmptyFirstPageHeaderMeansTheFirstPageHasNone()
    {
        Workbook book = Book(
            setup =>
            {
                setup.DifferentFirstPage = true;
                setup.Header = new HeaderFooterText { Default = "&CEvery page" };
            },
            rows: 120);

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain("Every page", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("Every page", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptyEvenPageHeaderMeansTheEvenPagesHaveNone()
    {
        Workbook book = Book(
            setup =>
            {
                setup.DifferentOddEven = true;
                setup.Header = new HeaderFooterText { Default = "&COdd side" };
            },
            rows: 120);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Odd side", rendered.Text(0), StringComparison.Ordinal);
        Assert.DoesNotContain("Odd side", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void APictureInAHeaderIsSaidToBeSkipped()
    {
        Workbook book = Book(setup => setup.Header = new HeaderFooterText { Default = "&L&G" });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.ContentSkipped && warning.Subject == "header-picture");
    }

    [Fact]
    public void APageNumberedFromASheetsOwnStartCountsFromThere()
    {
        Workbook book = Book(setup =>
        {
            setup.FirstPageNumber = 7;
            setup.Footer = new HeaderFooterText { Default = "&C&P" };
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("7", rendered.Text(0), StringComparison.Ordinal);
    }
}
