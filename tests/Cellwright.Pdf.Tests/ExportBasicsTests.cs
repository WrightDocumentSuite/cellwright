using Cellwright.Styles;
using Inkwright;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// The shape of a render: which sheets it prints, what paper it puts them on, and what the
/// workbook's own properties become.
/// </summary>
public sealed class ExportBasicsTests
{
    private static Workbook OneSheet(string text = "Hello")
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = text;
        return book;
    }

    [Fact]
    public void ASheetWithOneCellPrintsOnePage()
    {
        using Rendered rendered = Rendered.Of(OneSheet());

        Assert.Equal(1, rendered.PageCount);
        Assert.Contains("Hello", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptyWorkbookStillPrintsOnePage()
    {
        using Rendered rendered = Rendered.Of(Workbook.Create());

        Assert.Equal(1, rendered.PageCount);
    }

    [Fact]
    public void PrintingAWorkbookLeavesItAsItWasFound()
    {
        Workbook empty = Workbook.Create();
        using (Rendered.Of(empty))
        {
            Assert.Empty(empty.Sheets);
        }

        Workbook blank = Workbook.Create();
        blank.Sheets.Add("Sheet1");
        using (Rendered.Of(blank))
        {
            Assert.Single(blank.Sheets);
        }
    }

    [Fact]
    public void AnEmptySheetPrintedOnItsOwnGetsItsOwnPaper()
    {
        Workbook book = OneSheet("First");
        Worksheet second = book.Sheets.Add("Second");
        second.PageSetup.PaperSize = 9;
        second.PageSetup.Orientation = PageOrientation.Landscape;

        var options = new PdfExportOptions { Selection = SheetSelection.Named };
        options.Sheets.Add("Second");

        using Rendered rendered = Rendered.Of(book, options);

        Assert.Equal(1, rendered.PageCount);
        Assert.Equal(842, rendered.Document.Pages[0].Width, 1);
        Assert.Equal(595, rendered.Document.Pages[0].Height, 1);
    }

    [Fact]
    public void ThePaperIsTheOneThePageSetupNames()
    {
        Workbook book = OneSheet();
        book.Sheets[0].PageSetup.PaperSize = 9;

        using Rendered rendered = Rendered.Of(book);

        // A4 portrait, in points.
        Assert.Equal(595, rendered.Document.Pages[0].Width, 1);
        Assert.Equal(842, rendered.Document.Pages[0].Height, 1);
    }

    [Theory]
    [InlineData(3, 792, 1224)]      // Tabloid, eleven by seventeen inches
    [InlineData(4, 1224, 792)]      // Ledger, which is the same paper on its side
    [InlineData(14, 612, 936)]      // Folio, eight and a half by thirteen
    [InlineData(66, 1191, 1684)]    // A2
    [InlineData(70, 298, 420)]      // A6
    [InlineData(77, 842, 595)]      // A4 rotated
    public void EveryPaperTheFormatNamesIsTheSizeItSays(int code, double width, double height)
    {
        Workbook book = OneSheet();
        book.Sheets[0].PageSetup.PaperSize = code;

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(width, rendered.Document.Pages[0].Width, 1);
        Assert.Equal(height, rendered.Document.Pages[0].Height, 1);
        Assert.DoesNotContain(rendered.Diagnostics, warning => warning.Subject == $"paper-{code}");
    }

    [Fact]
    public void APaperCodeNobodyKnowsFallsBackToA4AndSaysSo()
    {
        Workbook book = OneSheet();
        book.Sheets[0].PageSetup.PaperSize = 250;

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(595, rendered.Document.Pages[0].Width, 1);
        Assert.Equal(842, rendered.Document.Pages[0].Height, 1);
        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.LayoutApproximated && warning.Subject == "paper-250");
    }

    [Fact]
    public void LandscapeTurnsThePaperOver()
    {
        Workbook book = OneSheet();
        book.Sheets[0].PageSetup.PaperSize = 9;
        book.Sheets[0].PageSetup.Orientation = PageOrientation.Landscape;

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.Document.Pages[0].Width > rendered.Document.Pages[0].Height);
    }

    [Fact]
    public void EverySheetWithContentGetsItsOwnPage()
    {
        Workbook book = OneSheet("First");
        book.Sheets.Add("Second")["A1"] = "Second";
        book.Sheets.Add("Third")["A1"] = "Third";

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(3, rendered.PageCount);
        Assert.Contains("Second", rendered.Text(1), StringComparison.Ordinal);
    }

    [Fact]
    public void AHiddenSheetIsNotPrintedUnlessAskedFor()
    {
        Workbook book = OneSheet("Shown");
        Worksheet hidden = book.Sheets.Add("Hidden");
        hidden["A1"] = "Secret";
        hidden.IsHidden = true;

        using Rendered visible = Rendered.Of(book);
        Assert.Equal(1, visible.PageCount);

        using Rendered all = Rendered.Of(book, new PdfExportOptions { Selection = SheetSelection.All });
        Assert.Equal(2, all.PageCount);
    }

    [Fact]
    public void OnlyTheSheetsNamedArePrintedWhenNamesAreGiven()
    {
        Workbook book = OneSheet("First");
        book.Sheets.Add("Wanted")["A1"] = "Wanted";
        book.Sheets.Add("Other")["A1"] = "Other";

        var options = new PdfExportOptions { Selection = SheetSelection.Named };
        options.Sheets.Add("Wanted");

        using Rendered rendered = Rendered.Of(book, options);

        Assert.Equal(1, rendered.PageCount);
        Assert.Contains("Wanted", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void ASheetNamedForPrintingThatIsNotThereIsSaidSo()
    {
        var options = new PdfExportOptions { Selection = SheetSelection.Named };
        options.Sheets.Add("Nowhere");

        using Rendered rendered = Rendered.Of(OneSheet(), options);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.ContentSkipped && warning.Subject == "Nowhere");
    }

    [Fact]
    public void TheWorkbooksPropertiesBecomeTheDocumentsOwn()
    {
        Workbook book = OneSheet();
        book.Properties.Title = "Quarterly report";
        book.Properties.Author = "The finance team";

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal("Quarterly report", rendered.Document.Info.Title);
        Assert.Equal("The finance team", rendered.Document.Info.Author);
    }

    [Fact]
    public void OneSheetCanBePrintedOnItsOwn()
    {
        Workbook book = OneSheet("First");
        Worksheet second = book.Sheets.Add("Second");
        second["A1"] = "Only this";

        PdfExportResult result = second.ToPdf();
        byte[] bytes;
        using (PdfDocument rendered = result.Document)
            bytes = rendered.ToArray();

        using PdfDocument pdf = PdfDocument.Load(bytes);

        Assert.Equal(1, result.PageCount);
        Assert.Contains("Only this", pdf.Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void AColumnIsAsWideAsItsDigitsSay()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "x";
        sheet["B1"] = "y";
        sheet.Column("A").SetWidth(20);

        using Rendered rendered = Rendered.Of(book);

        (double X, double Y)? first = rendered.Find("x");
        (double X, double Y)? second = rendered.Find("y");

        Assert.NotNull(first);
        Assert.NotNull(second);

        // Twenty digits of the standard eleven-point font come to a little over a hundred
        // points; the exact figure depends on the face, so the check is a generous band.
        double width = second!.Value.X - first!.Value.X;
        Assert.InRange(width, 90, 135);
    }

    [Fact]
    public void AMissingFontIsDrawnWithASubstituteAndSaidSo()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Text";
        sheet.Cell("A1").SetStyle(CellStyle.Default with
        {
            Font = FontStyle.Default with { Name = "No Such Face At All" },
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("Text", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains(rendered.Diagnostics.SubstitutedFonts, family => family == "No Such Face At All");
    }
}
