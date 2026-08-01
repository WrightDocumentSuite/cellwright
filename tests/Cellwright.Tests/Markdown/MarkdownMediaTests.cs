using System.Text;
using Xunit;

namespace Cellwright.Tests.Markdown;

/// <summary>
/// The picture half of a Markdown export: the gallery under a sheet's table and the files its
/// links point at.
/// </summary>
public sealed class MarkdownMediaTests
{
    /// <summary>
    /// A GFM table holds neither a spanning cell nor a floating object, so a picture cannot go
    /// where it sits on the sheet. It goes under the table, named by whatever the file called it.
    /// </summary>
    [Fact]
    public void PicturesBecomeAGalleryUnderTheTable()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Отчёт");
        sheet["A1"] = "Заголовок";
        sheet.Pictures.Add(Png(), "B2", "D6", "Логотип");

        MarkdownDocument document = book.ToMarkdown();

        Assert.Equal(Lf("""
            # Workbook

            ## Отчёт

            | Заголовок |
            | --- |

            ### Изображения

            ![Логотип](media/image1.png)

            """),
            document.Text);
        MarkdownImage image = Assert.Single(document.Images);
        Assert.Equal("image1.png", image.FileName);
        Assert.Equal(PictureFormat.Png, image.Format);
    }

    /// <summary>
    /// One logo on three sheets is one file and three links. The digest of the content decides,
    /// not the picture that asked for it.
    /// </summary>
    [Fact]
    public void IdenticalPicturesShareOneFile()
    {
        byte[] logo = Png();
        var book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        first["A1"] = "a";
        first.Pictures.Add(logo, "B2", "C3", "Logo");
        Worksheet second = book.Sheets.Add("Second");
        second["A1"] = "b";
        second.Pictures.Add(logo, "B2", "C3", "Logo again");
        second.Pictures.Add(Jpeg(), "D2", "E3", "Other");

        MarkdownDocument document = book.ToMarkdown();

        Assert.Equal(["image1.png", "image2.jpeg"], document.Images.Select(static i => i.FileName));
        Assert.Equal(2, CountOf(document.Text, "(media/image1.png)"));
        Assert.Equal(1, CountOf(document.Text, "(media/image2.jpeg)"));
    }

    /// <summary>The bytes are the picture's own: an export never re-encodes an image.</summary>
    [Fact]
    public void TheBytesAreHandedOverUntouched()
    {
        byte[] jpeg = Jpeg();
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("One");
        sheet["A1"] = "x";
        sheet.Pictures.Add(jpeg, "B2", "C3");

        Assert.Equal(jpeg, Assert.Single(book.ToMarkdown().Images).Content.ToArray());
    }

    /// <summary>Alternative text beats the shape's name, which is usually whatever Excel numbered it.</summary>
    [Fact]
    public void TheAlternativeTextIsPreferredForTheAltText()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Alt");
        sheet["A1"] = "x";
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.Between(CellRef.Parse("B2"), CellRef.Parse("C3")),
            Name = "Picture 1",
            Description = "Схема процесса",
        });

        Assert.Contains("![Схема процесса](media/image1.png)", book.ToMarkdown().Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// The gallery follows the anchors down the sheet rather than the drawing's stacking order:
    /// a document is read from the top, and painting order means nothing on a page.
    /// </summary>
    [Fact]
    public void TheGalleryFollowsTheAnchorsInReadingOrder()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Order");
        sheet["A1"] = "x";
        sheet.Pictures.Add(Png(), "A9", "B10", "нижняя");
        sheet.Pictures.Add(Jpeg(), "A2", "B3", "верхняя");

        string markdown = book.ToMarkdown().Text;

        Assert.True(
            markdown.IndexOf("верхняя", StringComparison.Ordinal) <
            markdown.IndexOf("нижняя", StringComparison.Ordinal));
    }

    /// <summary>A picture nobody sees on the sheet is not in the document either.</summary>
    [Fact]
    public void HiddenPicturesAreLeftOut()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Hidden");
        sheet["A1"] = "x";
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.Between(CellRef.Parse("B2"), CellRef.Parse("C3")),
            IsHidden = true,
        });

        MarkdownDocument document = book.ToMarkdown();

        Assert.Empty(document.Images);
        Assert.DoesNotContain("Изображения", document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void PicturesCanBeLeftOutAltogether()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("NoMedia");
        sheet["A1"] = "x";
        sheet.Pictures.Add(Png(), "B2", "C3", "Logo");

        MarkdownDocument document = book.ToMarkdown(new MarkdownExportOptions { IncludePictures = false });

        Assert.Empty(document.Images);
        Assert.DoesNotContain("image1.png", document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void TheMediaDirectoryCanBeRenamed()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Renamed");
        sheet["A1"] = "x";
        sheet.Pictures.Add(Png(), "B2", "C3", "Logo");

        MarkdownDocument document = book.ToMarkdown(new MarkdownExportOptions { MediaDirectoryName = "assets" });

        Assert.Contains("(assets/image1.png)", document.Text, StringComparison.Ordinal);
        Assert.Equal("assets", document.MediaDirectoryName);
    }

    [Theory]
    [InlineData("../outside")]
    [InlineData(@"..\outside")]
    [InlineData(".")]
    [InlineData("")]
    public void TheMediaDirectoryCannotLeaveOrReplaceTheExportDirectory(string directory)
    {
        var book = Workbook.Create();
        book.Sheets.Add("Unsafe")["A1"] = "x";

        Assert.Throws<ArgumentException>(() =>
            book.ToMarkdown(new MarkdownExportOptions { MediaDirectoryName = directory }));
    }

    [Fact]
    public void AnAbsoluteMediaDirectoryIsRejected()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Unsafe")["A1"] = "x";

        Assert.Throws<ArgumentException>(() =>
            book.ToMarkdown(new MarkdownExportOptions { MediaDirectoryName = Path.GetFullPath("outside") }));
    }

    [Fact]
    public async Task APortableMediaDirectoryIsNormalizedEncodedAndSaved()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Portable");
        sheet["A1"] = "x";
        sheet.Pictures.Add(Png(), "B2", "C3", "Logo");

        MarkdownDocument document = book.ToMarkdown(
            new MarkdownExportOptions { MediaDirectoryName = @"assets\quarter #1" });

        Assert.Equal("assets/quarter #1", document.MediaDirectoryName);
        Assert.Contains("(assets/quarter%20%231/image1.png)", document.Text, StringComparison.Ordinal);

        string directory = Path.Combine(Path.GetTempPath(), $"cellwright-md-{Guid.NewGuid():N}");
        try
        {
            await document.SaveAsync(directory, TestContext.Current.CancellationToken);
            Assert.True(File.Exists(Path.Combine(directory, "assets", "quarter #1", "image1.png")));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>The convenience layer: a directory with the document and its images beside it.</summary>
    [Fact]
    public async Task ExportWritesTheDocumentAndItsMedia()
    {
        byte[] logo = Png();
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Отчёт");
        sheet["A1"] = "Заголовок";
        sheet.Pictures.Add(logo, "B2", "D6", "Логотип");

        string directory = Path.Combine(Path.GetTempPath(), $"cellwright-md-{Guid.NewGuid():N}");
        try
        {
            await book.ExportMarkdownAsync(directory, cancellationToken: TestContext.Current.CancellationToken);

            string document = await File.ReadAllTextAsync(
                Path.Combine(directory, "workbook.md"), TestContext.Current.CancellationToken);
            Assert.Equal(book.ToMarkdown().Text, document);
            Assert.Equal(
                logo,
                await File.ReadAllBytesAsync(
                    Path.Combine(directory, "media", "image1.png"), TestContext.Current.CancellationToken));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ABookWithoutPicturesLeavesNoMediaDirectory()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Plain")["A1"] = "x";

        string directory = Path.Combine(Path.GetTempPath(), $"cellwright-md-{Guid.NewGuid():N}");
        try
        {
            await book.ExportMarkdownAsync(directory, cancellationToken: TestContext.Current.CancellationToken);

            Assert.True(File.Exists(Path.Combine(directory, "workbook.md")));
            Assert.False(Directory.Exists(Path.Combine(directory, "media")));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string Lf(string text) => text.ReplaceLineEndings("\n");

    private static int CountOf(string text, string value)
    {
        int count = 0;
        for (int i = text.IndexOf(value, StringComparison.Ordinal); i >= 0;
             i = text.IndexOf(value, i + value.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    /// <summary>The smallest valid PNG: a 1x1 fully transparent image.</summary>
    private static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>A minimal JPEG, enough to be recognized by its opening bytes.</summary>
    private static byte[] Jpeg()
    {
        var bytes = new List<byte> { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10 };
        bytes.AddRange(Encoding.ASCII.GetBytes("JFIF"));
        bytes.AddRange([0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xFF, 0xD9]);
        return [.. bytes];
    }
}
