using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Placing images on a sheet and getting them back out of the saved package.
/// </summary>
public sealed class PictureTests
{
    [Fact]
    public async Task APictureSurvivesTheRoundTrip()
    {
        byte[] image = Png();
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Pictures");
        sheet.Pictures.Add(image, "B2", "E8", "Logo");

        Workbook reloaded = await SaveAndLoadAsync(book);

        Picture picture = Assert.Single(reloaded.Sheets[0].Pictures);
        Assert.Equal(PictureFormat.Png, picture.Format);
        Assert.Equal(image, picture.Content.ToArray());
        Assert.Equal("Logo", picture.Name);
        Assert.Equal(PictureAnchorKind.TwoCell, picture.Anchor.Kind);
        Assert.Equal(CellRef.Parse("B2"), picture.Anchor.From);
        Assert.Equal(CellRef.Parse("E8"), picture.Anchor.To);
    }

    /// <summary>
    /// All three ways of pinning a picture have to come back as they went in. They are not
    /// interchangeable: one resizes with the cells, one moves without resizing, and one ignores
    /// the grid entirely.
    /// </summary>
    [Fact]
    public async Task EveryKindOfAnchorComesBackAsItself()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Anchors");
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.Between(CellRef.Parse("A1"), CellRef.Parse("C4")),
        });
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = PictureAnchor.At(CellRef.Parse("E5"), Emu.FromPixels(120), Emu.FromPixels(60)),
        });
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.Absolute,
                Position = (Emu.PerInch, Emu.PerInch * 2),
                Extent = (Emu.PerInch, Emu.PerInch),
            },
        });

        PictureCollection pictures = (await SaveAndLoadAsync(book)).Sheets[0].Pictures;

        Assert.Equal(3, pictures.Count);
        Assert.Equal(PictureAnchorKind.TwoCell, pictures[0].Anchor.Kind);
        Assert.Equal(CellRef.Parse("C4"), pictures[0].Anchor.To);

        Assert.Equal(PictureAnchorKind.OneCell, pictures[1].Anchor.Kind);
        Assert.Equal(CellRef.Parse("E5"), pictures[1].Anchor.From);
        Assert.Equal((Emu.FromPixels(120), Emu.FromPixels(60)), pictures[1].Anchor.Extent);

        Assert.Equal(PictureAnchorKind.Absolute, pictures[2].Anchor.Kind);
        Assert.Equal((Emu.PerInch, Emu.PerInch * 2), pictures[2].Anchor.Position);
    }

    /// <summary>
    /// The bytes are handed back exactly as they arrived. Re-encoding an image would be a
    /// silent quality loss, so the library never decodes one.
    /// </summary>
    [Fact]
    public async Task TheImageIsStoredByteForByte()
    {
        byte[] image = Jpeg();
        var book = Workbook.Create();
        book.Sheets.Add("One").Pictures.Add(image, "A1", "B2");

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
        ZipArchiveEntry entry = Assert.Single(
            zip.Entries, static e => e.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
        Assert.Equal("xl/media/image1.jpeg", entry.FullName);

        using var content = new MemoryStream();
        entry.Open().CopyTo(content);
        Assert.Equal(image, content.ToArray());
    }

    [Fact]
    public async Task TheWrittenPackageIsValid()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Pictures");
        sheet["A1"] = "beside a picture";
        sheet.Pictures.Add(Png(), "B2", "D6", "Logo");
        sheet.Comments.Add("A1", "and a note");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        ValidationErrorInfo? error = new OpenXmlValidator()
            .Validate(document, TestContext.Current.CancellationToken)
            .FirstOrDefault();
        Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
    }

    /// <summary>
    /// Two sheets showing the same image still get one media part each. Sharing the part would
    /// be smaller, but it would also mean deleting one sheet's picture could blank the other's.
    /// </summary>
    [Fact]
    public async Task EachSheetKeepsItsOwnPictures()
    {
        var book = Workbook.Create();
        book.Sheets.Add("First").Pictures.Add(Png(), "A1", "B2");
        book.Sheets.Add("Second").Pictures.Add(Jpeg(), "C3", "D4");

        Workbook reloaded = await SaveAndLoadAsync(book);

        Assert.Equal(PictureFormat.Png, Assert.Single(reloaded.Sheets[0].Pictures).Format);
        Assert.Equal(PictureFormat.Jpeg, Assert.Single(reloaded.Sheets[1].Pictures).Format);
    }

    /// <summary>
    /// A picture whose cells move with an inserted row moves too, and one pinned to the sheet
    /// does not. Both are what Excel does, and the difference is the whole point of the flag.
    /// </summary>
    [Fact]
    public void InsertingRowsMovesThePicturesAboveWhichItHappens()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Shifted");
        Picture moving = sheet.Pictures.Add(Png(), "B5", "D9");
        Picture pinned = sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.TwoCell,
                From = CellRef.Parse("F5"),
                To = CellRef.Parse("H9"),
                MoveBehavior = PictureMoveBehavior.Fixed,
            },
        });

        sheet.InsertRows(2, 3);

        Assert.Equal(CellRef.Parse("B8"), moving.Anchor.From);
        Assert.Equal(CellRef.Parse("D12"), moving.Anchor.To);
        Assert.Equal(CellRef.Parse("F5"), pinned.Anchor.From);
    }

    /// <summary>
    /// A sheet whose pictures were never touched keeps the drawing part it arrived with, along
    /// with whatever else was in it. Rewriting it from the model would cost the charts.
    /// </summary>
    [Fact]
    public async Task AnUntouchedDrawingIsNotRewritten()
    {
        var book = Workbook.Create();
        book.Sheets.Add("Empty")["A1"] = 1;

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
        Assert.DoesNotContain(zip.Entries, static e => e.FullName.StartsWith("xl/drawings/", StringComparison.Ordinal));
    }

    private static async Task<Workbook> SaveAndLoadAsync(Workbook book)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(
            stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        return await Workbook.LoadAsync(stream, cancellationToken: TestContext.Current.CancellationToken);
    }

    /// <summary>The smallest valid PNG: an 1x1 fully transparent image.</summary>
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
