using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// By default a render carries on past what it cannot draw and says so afterwards. A caller can
/// instead name the compromises that must stop it, and then gets no document at all.
/// </summary>
public sealed class StrictExportTests
{
    private static Workbook WithComment()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Total";
        sheet.Comments.Add(CellRef.Parse("A1"), "Check this", "Reviewer");
        return book;
    }

    [Fact]
    public void ByDefaultARenderCarriesOnPastWhatItCannotDraw()
    {
        using Rendered rendered = Rendered.Of(WithComment());

        Assert.Contains(rendered.Diagnostics, static warning => warning.Kind == PdfExportWarningKind.ContentSkipped);
    }

    [Fact]
    public void ARefusedOmissionStopsTheRenderAndNamesWhatWasRefused()
    {
        var options = new PdfExportOptions { RefuseWhen = PdfExportRefusal.Omissions };

        PdfExportRefusedException exception = Assert.Throws<PdfExportRefusedException>(
            () => PdfExporter.Render(WithComment(), options));

        Assert.Contains(exception.Refused, static warning => warning.Subject == "comment");
        Assert.Equal(exception.Refused[0], exception.Diagnostics.First(warning => warning == exception.Refused[0]));
    }

    [Fact]
    public void ARefusedRenderCreatesNoFile()
    {
        string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.pdf");
        var options = new PdfExportOptions { RefuseWhen = PdfExportRefusal.Omissions };

        Assert.Throws<PdfExportRefusedException>(() => PdfExporter.Save(WithComment(), path, options));

        Assert.False(File.Exists(path));
    }

    [Fact]
    public void AWorkbookWithNoRefusedCompromiseIsProduced()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = "Total";
        var options = new PdfExportOptions { RefuseWhen = PdfExportRefusal.Omissions };

        using Rendered rendered = Rendered.Of(book, options);

        Assert.True(rendered.Document.Pages.Count > 0);
    }

    [Fact]
    public void APolicyCanLookAtTheSubject()
    {
        var acceptsComments = new PdfExportOptions { RefuseWhen = static warning => warning.Subject == "picture" };
        using (Rendered.Of(WithComment(), acceptsComments))
        {
        }

        var refusesComments = new PdfExportOptions { RefuseWhen = static warning => warning.Subject == "comment" };
        Assert.Throws<PdfExportRefusedException>(() => PdfExporter.Render(WithComment(), refusesComments));
    }
}
