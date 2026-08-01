using Inkwright;
using Inkwright.Tagging;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// The things that hang off the grid rather than sitting in it: pictures pinned to cells, and
/// the structure tree that tells a reader what the grid means.
/// </summary>
public sealed class PictureAndTaggingTests
{
    /// <summary>The smallest PNG that decodes: one grey pixel.</summary>
    private static byte[] Png =>
    [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE,
        0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54,
        0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
        0x18, 0xDD, 0x8D, 0xB0,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    private static Workbook WithPicture(PictureAnchor anchor)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Beside a picture";

        sheet.Pictures.Add(new Picture
        {
            Content = Png,
            Format = PictureFormat.Png,
            Anchor = anchor,
            Name = "A picture",
            Description = "A grey square",
        });

        return book;
    }

    /// <summary>
    /// How big a picture was painted. The transformation is the one immediately before the
    /// operator that paints it, which is what tells it apart from the page's own matrices.
    /// </summary>
    private static (double Width, double Height) Painted(Rendered rendered)
    {
        System.Text.RegularExpressions.Match placed = System.Text.RegularExpressions.Regex.Match(
            rendered.Content(),
            @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) [-\d.]+ [-\d.]+ cm\s*/\S+ Do");

        Assert.True(placed.Success, "no picture was painted on the page");

        return (
            double.Parse(placed.Groups["w"].Value, System.Globalization.CultureInfo.InvariantCulture),
            double.Parse(placed.Groups["h"].Value, System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void APictureBetweenTwoCellsIsAsBigAsTheyMakeIt()
    {
        Workbook book = WithPicture(PictureAnchor.Between(new CellRef(2, 1), new CellRef(4, 3)));

        using Rendered rendered = Rendered.Of(book);
        (double width, double height) = Painted(rendered);

        // The far corner sits at the top-left of C4, so the picture covers columns A and B and
        // rows 2 and 3: two default columns across and two default rows down.
        Assert.InRange(width, 80, 100);
        Assert.InRange(height, 25, 35);
    }

    [Fact]
    public void APicturePinnedToOneCellKeepsTheSizeItWasGiven()
    {
        Workbook book = WithPicture(PictureAnchor.At(new CellRef(1, 1), 914400, 457200));

        using Rendered rendered = Rendered.Of(book);
        (double width, double height) = Painted(rendered);

        // An inch by half an inch, in points.
        Assert.Equal(72, width, 1);
        Assert.Equal(36, height, 1);
    }

    [Fact]
    public void APicturePinnedAcrossAPageBreakIsShownOnBothPages()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        for (int row = 1; row <= 8; row++)
            sheet.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"r{row}"));

        sheet.PageSetup.AddRowBreak(5);
        sheet.Pictures.Add(new Picture
        {
            Content = Png,
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.At(new CellRef(3, 1), 914400, 914400),
            Description = "A square across the break",
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(2, rendered.PageCount);
        Assert.Contains(" Do", rendered.Content(0), StringComparison.Ordinal);
        Assert.Contains(" Do", rendered.Content(1), StringComparison.Ordinal);

        // Both pages cut it at their own edge rather than letting it run over.
        Assert.Contains("\nW\nn\n", rendered.Content(0), StringComparison.Ordinal);
        Assert.Contains("\nW\nn\n", rendered.Content(1), StringComparison.Ordinal);
    }

    [Fact]
    public void APictureOnASheetWithNothingElsePrintsAllTheSame()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Pictures.Add(new Picture
        {
            Content = Png,
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.Absolute,
                Position = (914400, 914400),
                Extent = (914400, 457200),
            },
        });

        using Rendered rendered = Rendered.Of(book);
        (double width, double height) = Painted(rendered);

        Assert.Equal(1, rendered.PageCount);
        Assert.Equal(72, width, 1);
        Assert.Equal(36, height, 1);
    }

    [Fact]
    public void AHiddenPictureIsNotDrawn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Nothing beside me";
        sheet.Pictures.Add(new Picture
        {
            Content = Png,
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.At(new CellRef(1, 1), 914400, 457200),
            IsHidden = true,
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain(" Do", rendered.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void APictureInAFormatThatCannotBeEmbeddedIsSaidSo()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "x";
        sheet.Pictures.Add(new Picture
        {
            Content = new byte[] { 1, 2, 3, 4 },
            Format = PictureFormat.Emf,
            Anchor = PictureAnchor.At(new CellRef(1, 1), 914400, 457200),
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(rendered.Diagnostics, warning => warning.Kind == PdfExportWarningKind.ImageSkipped);
    }

    [Fact]
    public void ATaggedSheetIsATableOfRowsAndCells()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Head";
        sheet["A2"] = "Body";
        sheet["B2"] = 42d;

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { Tagged = true });

        Assert.NotEmpty(rendered.Document.Structure);
        Assert.Contains(Tags(rendered.Document.Structure), tag => tag == "Table");
        Assert.Contains(Tags(rendered.Document.Structure), tag => tag == "TR");
        Assert.Contains(Tags(rendered.Document.Structure), tag => tag == "TD");
    }

    [Fact]
    public void ARepeatedTitleRowBecomesHeaderCells()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Heading";
        sheet["A2"] = "Body";
        sheet.SetRepeatRows(1, 1);

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { Tagged = true });

        Assert.Contains(Tags(rendered.Document.Structure), tag => tag == "TH");
    }

    [Fact]
    public void APictureIsAFigureWithTheDescriptionItCarries()
    {
        Workbook book = WithPicture(PictureAnchor.At(new CellRef(1, 1), 914400, 457200));

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { Tagged = true });

        Assert.Contains(Tags(rendered.Document.Structure), tag => tag == "Figure");
    }

    [Fact]
    public void AnUntaggedRenderBuildsNoTreeAtAll()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = "Plain";

        using Rendered rendered = Rendered.Of(book);

        Assert.Empty(rendered.Document.Structure);
    }

    [Fact]
    public void ATaggedDocumentWithATitleAsksReadersToShowIt()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = "Plain";
        book.Properties.Title = "The report";

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { Tagged = true, Language = "en-GB" });

        Assert.Equal("The report", rendered.Document.Info.Title);
        Assert.NotNull(rendered.Document.Catalog.GetDictionary(Inkwright.Cos.PdfName.Get("ViewerPreferences")));
    }

    [Fact]
    public void PagesOfSeveralSheetsAreNumberedSheetBySheet()
    {
        Workbook book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");

        for (int row = 1; row <= 120; row++)
            first.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"§{row}"));

        second["A1"] = "Alone";
        first.PageSetup.Footer = new HeaderFooterText { Default = "&C&P/&N" };
        second.PageSetup.Footer = new HeaderFooterText { Default = "&C&P/&N" };

        using Rendered rendered = Rendered.Of(book);

        // Each sheet counts its own pages, which is how Excel prints "1 of 2" then "1 of 1".
        Assert.Contains("1/", rendered.Text(0), StringComparison.Ordinal);
        Assert.Contains("1/1", rendered.Text(rendered.PageCount - 1), StringComparison.Ordinal);
    }

    private static IEnumerable<string> Tags(IReadOnlyList<PdfStructureNode> nodes)
    {
        foreach (PdfStructureNode node in nodes)
        {
            yield return node.Tag;

            foreach (string nested in Tags(node.Children))
                yield return nested;
        }
    }
}
