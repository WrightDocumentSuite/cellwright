using Cellwright.Styles;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// The promise the diagnostics make: everything the render leaves out is said, so that a caller
/// comparing a PDF with the workbook it came from never has to guess what became of something.
/// </summary>
public sealed class DiagnosticsTests
{
    private static Worksheet Sheet(Workbook book)
    {
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Something to print";
        return sheet;
    }

    [Fact]
    public void APlainSheetHasNothingToOwnUpTo()
    {
        Workbook book = Workbook.Create();
        Sheet(book);

        using Rendered rendered = Rendered.Of(book);

        Assert.True(rendered.Diagnostics.IsEmpty, rendered.Diagnostics.ToString());
    }

    [Fact]
    public void ACommentIsSaidToBeLeftOut()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet.Comments.Add(CellRef.Parse("A1"), "Check this", "Reviewer");

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.ContentSkipped && warning.Subject == "comment");
    }

    [Fact]
    public void AThreadedCommentIsSaidToBeLeftOut()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        Person person = book.People.Add("Reviewer");
        ThreadedComment thread = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Check this", person);
        sheet.ThreadedComments.Reply(thread, "Agreed", person);

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.ContentSkipped &&
                warning.Subject == "threaded-comments");
    }

    [Fact]
    public void AConditionalFormatThatWasWorkedOutIsNotComplainedAbout()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = 7d;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 5, new DifferentialStyle { Bold = true }));

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("conditional-format", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void AConditionalFormatNeedingAFormulaSaysSoRatherThanGuessing()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = 7d;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression(
            "A1:A10", "$A1>5", new DifferentialStyle { Bold = true }));

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.LayoutApproximated
                && warning.Subject == "conditional-format-formula");
    }

    [Fact]
    public void PrintedHeadingsAreDrawnWithoutDiagnostic()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["B2"] = "Cell value";
        sheet.PageSetup.PrintHeadings = true;

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject == "print-headings");
        Assert.Contains("A", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("B", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("1", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("2", rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void DraftAndBlackAndWhiteAreSaidToBeIgnored()
    {
        Workbook book = Workbook.Create();
        PageSetup setup = Sheet(book).PageSetup;
        setup.Draft = true;
        setup.BlackAndWhite = true;

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "draft");
        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "black-and-white");
    }

    [Fact]
    public void AHeaderOnAScaledSheetIsSaidNotToBeScaledWithIt()
    {
        Workbook book = Workbook.Create();
        PageSetup setup = Sheet(book).PageSetup;
        setup.Scale = 50;
        setup.Header = new HeaderFooterText { Default = "&CReport" };

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.LayoutApproximated && warning.Subject == "header-scale");
    }

    [Fact]
    public void AHeaderThatDoesNotFollowTheMarginsIsSaidToAllTheSame()
    {
        Workbook book = Workbook.Create();
        PageSetup setup = Sheet(book).PageSetup;
        setup.AlignHeaderWithMargins = false;
        setup.Header = new HeaderFooterText { Default = "&CReport" };

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Kind == PdfExportWarningKind.LayoutApproximated && warning.Subject == "header-margins");
    }

    [Fact]
    public void TheSameComplaintAboutTheSameThingIsMadeOnlyOnce()
    {
        Workbook book = Workbook.Create();

        for (int i = 1; i <= 3; i++)
        {
            Worksheet sheet = book.Sheets.Add($"Sheet{i}");
            sheet["A1"] = "Something to print";
            sheet.PageSetup.Draft = true;
        }

        using Rendered rendered = Rendered.Of(book);

        Assert.Single(rendered.Diagnostics, warning => warning.Subject == "draft");
    }
}
