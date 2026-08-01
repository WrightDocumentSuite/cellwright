using System.Text;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Writing pictures to .xls and reading them back: the image store, the shapes that point into
/// it, and the anchors that place them. The reader is the oracle — it decodes the Escher layer
/// independently, so a shape it recognises is a shape any reader of the classic records will.
/// </summary>
public sealed class PictureWriteTests
{
    private static async Task<Worksheet> RoundTripAsync(Workbook book, List<WorkbookWarning>? warnings = null)
    {
        using var stream = new MemoryStream();
        await XlsWorkbook.SaveAsync(
            book,
            stream,
            new XlsSaveOptions { OnWarning = warnings is null ? null : warnings.Add },
            TestContext.Current.CancellationToken);

        stream.Position = 0;
        Workbook read = await XlsWorkbook.LoadAsync(stream, options: null, TestContext.Current.CancellationToken);
        return read.Sheets[0];
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

    [Fact]
    public async Task APictureSurvivesTheRoundTrip()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "beside the picture";
        sheet.Pictures.Add(Png(), "B2", "D6", "Logo");

        Worksheet read = await RoundTripAsync(book);

        Picture picture = Assert.Single(read.Pictures);
        Assert.Equal(Png(), picture.Content.ToArray());
        Assert.Equal(PictureFormat.Png, picture.Format);
    }

    [Fact]
    public async Task TheAnchorCellsComeBackWhereTheyWere()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Pictures.Add(Png(), "B2", "D6");

        Worksheet read = await RoundTripAsync(book);
        PictureAnchor anchor = Assert.Single(read.Pictures).Anchor;

        Assert.Equal(PictureAnchorKind.TwoCell, anchor.Kind);
        Assert.Equal(CellRef.Parse("B2"), anchor.From);
        Assert.Equal(CellRef.Parse("D6"), anchor.To);
    }

    [Fact]
    public async Task AnOffsetIntoACellSurvivesToTheFractionTheFormatKeeps()
    {
        // Offsets travel as 1024ths of a column and 256ths of a row, so what comes back is the
        // nearest such fraction rather than the exact EMU — half a default cell stays half.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        long halfColumn = (long)(61 * Emu.PerPixel) / 2;    // Half of the default 8.43-character column.
        long halfRow = (long)(15 * Emu.PerPoint) / 2;

        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.TwoCell,
                From = CellRef.Parse("B2"),
                FromOffset = (halfColumn, halfRow),
                To = CellRef.Parse("D6"),
            },
        });

        Worksheet read = await RoundTripAsync(book);
        PictureAnchor anchor = Assert.Single(read.Pictures).Anchor;

        Assert.InRange(anchor.FromOffset.X, halfColumn - Emu.PerPixel, halfColumn + Emu.PerPixel);
        Assert.InRange(anchor.FromOffset.Y, halfRow - Emu.PerPoint, halfRow + Emu.PerPoint);
    }

    [Fact]
    public async Task AOneCellAnchorIsWalkedOutToItsFarCorner()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        // Two default columns and four default rows of extent, from B2: the far corner lands at D6.
        long width = 2 * 64 * Emu.PerPixel;    // A default column is round(8.43 * 7 + 5) = 64 pixels.
        long height = (long)(4 * 15 * Emu.PerPoint);

        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.OneCell,
                From = CellRef.Parse("B2"),
                Extent = (width, height),
            },
        });

        Worksheet read = await RoundTripAsync(book);
        PictureAnchor anchor = Assert.Single(read.Pictures).Anchor;

        Assert.Equal(CellRef.Parse("B2"), anchor.From);
        Assert.Equal(CellRef.Parse("D6"), anchor.To);
    }

    [Fact]
    public async Task TwoPicturesOfTheSameImageShareOneStoreEntry()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Pictures.Add(Png(), "A1", "B2");
        sheet.Pictures.Add(Png(), "D4", "E5");

        using var stream = new MemoryStream();
        await XlsWorkbook.SaveAsync(book, stream, options: null, TestContext.Current.CancellationToken);

        // One PNG in the whole file: the second shape points at the first entry.
        byte[] bytes = stream.ToArray();
        byte[] signature = [0x89, 0x50, 0x4E, 0x47];
        int occurrences = 0;
        for (int i = 0; i + signature.Length <= bytes.Length; i++)
        {
            if (bytes.AsSpan(i, signature.Length).SequenceEqual(signature))
                occurrences++;
        }

        Assert.Equal(1, occurrences);

        stream.Position = 0;
        Workbook read = await XlsWorkbook.LoadAsync(stream, options: null, TestContext.Current.CancellationToken);
        Assert.Equal(2, read.Sheets[0].Pictures.Count);
        Assert.Equal(Png(), read.Sheets[0].Pictures[1].Content.ToArray());
    }

    [Fact]
    public async Task DifferentFormatsLiveSideBySide()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Pictures.Add(Png(), "A1", "B2");
        sheet.Pictures.Add(Jpeg(), "D4", "E5");

        Worksheet read = await RoundTripAsync(book);

        Assert.Equal(2, read.Pictures.Count);
        Assert.Equal(PictureFormat.Png, read.Pictures[0].Format);
        Assert.Equal(PictureFormat.Jpeg, read.Pictures[1].Format);
        Assert.Equal(Jpeg(), read.Pictures[1].Content.ToArray());
    }

    [Fact]
    public async Task PicturesAndNotesShareOneDrawing()
    {
        // The real point of one writer: both kinds under one group shape, one id block, and a
        // file Excel does not send to repair over a duplicated identifier.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Pictures.Add(Png(), "B2", "D6");
        sheet.Comments.Add(CellRef.Parse("A1"), "A note beside a picture", "Reviewer");

        Worksheet read = await RoundTripAsync(book);

        Assert.Single(read.Pictures);
        CellComment note = Assert.Single(read.Comments);
        Assert.Equal("A note beside a picture", note.Text);
    }

    [Fact]
    public async Task AFormatTheStoreCannotHoldIsReportedNotDropped()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        byte[] gif = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0];
        sheet.Pictures.Add(new Picture { Content = gif, Format = PictureFormat.Gif, Anchor = PictureAnchor.Between(CellRef.Parse("A1"), CellRef.Parse("B2")) });

        var warnings = new List<WorkbookWarning>();
        Worksheet read = await RoundTripAsync(book, warnings);

        Assert.Empty(read.Pictures);
        Assert.Contains(warnings, w => w.Code == "PictureFormatNotWritten");
    }

    [Fact]
    public async Task PicturesOnTwoSheetsClaimSeparateIdBlocks()
    {
        Workbook book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");
        first.Pictures.Add(Png(), "A1", "B2");
        second.Pictures.Add(Jpeg(), "C3", "D4");

        using var stream = new MemoryStream();
        await XlsWorkbook.SaveAsync(book, stream, options: null, TestContext.Current.CancellationToken);
        stream.Position = 0;
        Workbook read = await XlsWorkbook.LoadAsync(stream, options: null, TestContext.Current.CancellationToken);

        Assert.Single(read.Sheets[0].Pictures);
        Assert.Single(read.Sheets[1].Pictures);
        Assert.Equal(PictureFormat.Jpeg, read.Sheets[1].Pictures[0].Format);
    }

    [Fact]
    public async Task AWidenedColumnChangesTheFractionButNotThePlace()
    {
        // The offset is a fraction of the column it falls in, so the same EMU offset into a
        // wider column is a smaller fraction — and must still come back as the same EMU.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Column("B").SetWidth(30);

        long offset = 50 * Emu.PerPixel;
        sheet.Pictures.Add(new Picture
        {
            Content = Png(),
            Format = PictureFormat.Png,
            Anchor = new PictureAnchor
            {
                Kind = PictureAnchorKind.TwoCell,
                From = CellRef.Parse("B2"),
                FromOffset = (offset, 0),
                To = CellRef.Parse("D6"),
            },
        });

        Worksheet read = await RoundTripAsync(book);

        Assert.InRange(
            Assert.Single(read.Pictures).Anchor.FromOffset.X,
            offset - Emu.PerPixel, offset + Emu.PerPixel);
    }
}
