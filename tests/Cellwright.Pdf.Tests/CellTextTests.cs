using Cellwright.Styles;
using Inkwright.Text;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// What happens to a cell's text when it does not simply fit: wrapped, spilled, clipped, shrunk
/// or turned, and coloured by the format that rendered it.
/// </summary>
public sealed class CellTextTests
{
    private const string Long = "A sentence that is far too long for the column it sits in";

    private static Worksheet Sheet(Workbook book) => book.Sheets.Add("Sheet1");

    private static CellStyle Wrapped => CellStyle.Default with
    {
        Alignment = Alignment.Default with { WrapText = true },
    };

    [Fact]
    public void WrappedTextGoesOntoSeveralLines()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = Long;
        sheet.Cell("A1").SetStyle(Wrapped);
        sheet.Column("A").SetWidth(12);

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.Lines().Count >= 4, $"the text came out on {rendered.Lines().Count} lines");
    }

    [Fact]
    public void AWrappedRowGrowsToHoldItsLines()
    {
        Workbook tall = Workbook.Create();
        Worksheet sheet = Sheet(tall);
        sheet["A1"] = Long;
        sheet.Cell("A1").SetStyle(Wrapped);
        sheet.Column("A").SetWidth(12);
        sheet["A2"] = "Below";

        using Rendered rendered = Rendered.Of(tall);

        (double X, double Y)? first = rendered.Find("A");
        (double X, double Y)? below = rendered.Find("B");

        Assert.NotNull(first);
        Assert.NotNull(below);
        Assert.True(
            first!.Value.Y - below!.Value.Y > 40,
            "the row did not grow to hold the wrapped lines");
    }

    [Fact]
    public void ARowThatStatesItsHeightIsLeftAtIt()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = Long;
        sheet.Cell("A1").SetStyle(Wrapped);
        sheet.Column("A").SetWidth(12);
        sheet.Row(1).SetHeight(15);
        sheet["A2"] = "Below";

        using Rendered rendered = Rendered.Of(book);

        (double X, double Y)? below = rendered.Find("B");
        Assert.NotNull(below);

        // The row kept its fifteen points, so the row under it starts near the top of the page
        // rather than five wrapped lines down. A4 is 842 points tall and the margin is 54.
        double fromTop = 842 - below!.Value.Y;
        Assert.InRange(fromTop, 54, 100);

        // Only what fits the short row is drawn: bottom-aligned, that is the last wrapped line.
        Assert.Equal(2, rendered.Lines().Count);
        Assert.DoesNotContain("A sentence", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void TextSpillsIntoTheEmptyCellBesideIt()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = Long;
        sheet.Column("A").SetWidth(8);

        using Rendered rendered = Rendered.Of(book);

        // Nothing is in B1 or C1, so the whole sentence is drawn.
        Assert.Contains("sits in", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void TextIsCutOffAtTheFirstNeighbourThatHoldsSomething()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = Long;
        sheet["B1"] = "Stop";
        sheet.Column("A").SetWidth(8);

        using Rendered rendered = Rendered.Of(book);

        // The neighbour is drawn and the sentence stops short of it; the two interleave in
        // extracted text, so the check is on the letters rather than on the joined line.
        Assert.NotNull(rendered.Find("S"));
        Assert.DoesNotContain("sits in", rendered.Text(0), StringComparison.Ordinal);
        Assert.DoesNotContain("column", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void ShrinkToFitMakesTheTextSmallerRatherThanSpill()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = Long;
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { ShrinkToFit = true },
        });
        sheet.Column("A").SetWidth(12);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> letters = rendered.Letters();

        Assert.NotEmpty(letters);
        Assert.Contains("sits in", rendered.Text(0), StringComparison.Ordinal);

        // Everything is on one line, and that line ends inside the column.
        Assert.Single(rendered.Lines());
        Assert.True(letters.Max(letter => letter.Origin.X + letter.Width) < 130);
    }

    [Fact]
    public void AnIndentMovesTheTextInFromTheEdge()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "x";
        sheet["A2"] = "x";
        sheet.Cell("A2").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { Indent = 2 },
        });
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> letters = [.. rendered.Letters().Where(letter => letter.Text == "x")];

        Assert.Equal(2, letters.Count);
        Assert.True(letters.Max(letter => letter.Origin.X) - letters.Min(letter => letter.Origin.X) > 10);
    }

    [Fact]
    public void ANegativeNumberTakesTheColourItsFormatNames()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = -5d;
        sheet.Cell("A1").SetStyle(CellStyle.Default with { NumberFormat = "#,##0;[Red]-#,##0" });
        sheet.Column("A").SetWidth(12);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("1 0 0 rg", rendered.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void TurnedTextIsDrawnUnderAMatrix()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Turned";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { TextRotation = 90 },
        });
        sheet.Row(1).SetHeight(60);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> letters = rendered.Letters();

        Assert.NotEmpty(letters);

        // Turned text stacks up the page rather than across it.
        double spreadX = letters.Max(l => l.Origin.X) - letters.Min(l => l.Origin.X);
        double spreadY = letters.Max(l => l.Origin.Y) - letters.Min(l => l.Origin.Y);
        Assert.True(spreadY > spreadX, $"the text was not turned: {spreadX:0.0} across, {spreadY:0.0} down");
    }

    [Fact]
    public void StackedTextPutsOneLetterUnderAnother()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "ABC";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Alignment = Alignment.Default with { TextRotation = Alignment.StackedRotation },
        });
        sheet.Row(1).SetHeight(60);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> letters = rendered.Letters();

        Assert.Equal(3, letters.Count);
        Assert.Equal(3, letters.Select(letter => Math.Round(letter.Origin.Y)).Distinct().Count());
    }

    [Fact]
    public void HebrewInACellIsDrawnFromTheRight()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "\u05E9\u05DC\u05D5\u05DD";
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);
        IReadOnlyList<PdfLetter> letters = [.. rendered.Letters().Where(l => l.Text.Length == 1 && l.Text[0] >= '\u0590')];

        Assert.NotEmpty(letters);

        // The first letter read — shin — is the rightmost letter drawn.
        Assert.Equal("\u05E9", letters.OrderByDescending(letter => letter.Origin.X).First().Text);
    }

    [Fact]
    public void AHyperlinkBecomesAClickableArea()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Follow me";
        sheet.Hyperlinks.Add(Hyperlink.ToUrl(RangeRef.Parse("A1"), "https://example.org/"));

        using Rendered rendered = Rendered.Of(book);

        Assert.NotEmpty(rendered.Document.Pages[0].Annotations);
        AssertClickable(rendered, "Follow me");
    }

    [Fact]
    public void AClickableAreaFollowsItsCellWhenTheSheetIsScaled()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Follow me";
        sheet.Hyperlinks.Add(Hyperlink.ToUrl(RangeRef.Parse("A1"), "https://example.org/"));
        sheet.PageSetup.Scale = 60;

        using Rendered rendered = Rendered.Of(book);

        AssertClickable(rendered, "Follow me");
    }

    [Fact]
    public void AClickableAreaFollowsItsCellWhenTheGridIsCentred()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Follow me";
        sheet.Hyperlinks.Add(Hyperlink.ToUrl(RangeRef.Parse("A1"), "https://example.org/"));
        sheet.PageSetup.CenterHorizontally = true;
        sheet.PageSetup.CenterVertically = true;

        using Rendered rendered = Rendered.Of(book);

        AssertClickable(rendered, "Follow me");
    }

    /// <summary>
    /// The clickable area has to sit over the text it belongs to. An annotation is read in page
    /// space, so nothing the content stream does to the grid reaches it: the scale and the place
    /// the grid was put on the paper have to be in the rectangle already.
    /// </summary>
    private static void AssertClickable(Rendered rendered, string text)
    {
        Inkwright.PdfRectangle area = rendered.Document.Pages[0].Annotations[0].Rectangle;
        (double X, double Y) at = rendered.Find(text)
            ?? throw new Xunit.Sdk.XunitException($"'{text}' was not drawn at all");

        Assert.True(
            at.X >= area.Left && at.X <= area.Right && at.Y >= area.Bottom && at.Y <= area.Top,
            $"the text at ({at.X:0.#}, {at.Y:0.#}) is outside the clickable area " +
            $"[{area.Left:0.#}…{area.Right:0.#}, {area.Bottom:0.#}…{area.Top:0.#}]");
    }

    [Fact]
    public void ALinkInsideTheWorkbookIsSaidToBeInert()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Elsewhere";
        sheet.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("A1"), "Sheet1!B2"));

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.ContentSkipped && warning.Subject == "internal-link");
    }

    [Fact]
    public void AnUnderlineIsRuledUnderTheText()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "Underlined";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Font = FontStyle.Default with { Underline = UnderlineStyle.Single },
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("\nS\n", rendered.Content(), StringComparison.Ordinal);
    }
}
