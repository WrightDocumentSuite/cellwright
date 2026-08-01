using System.Buffers.Binary;
using System.IO.Compression;
using Cellwright.Xls.Reading;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The corners of the Office Art layer the corpus does not reach, on trees built here byte by
/// byte from MS-ODRAW.
/// </summary>
/// <remarks>
/// A fixture assembled here shows only that the reader agrees with this file's reading of the
/// specification, which is weaker evidence than a file Excel wrote. The corpus comparison
/// against NPOI is what carries the weight; these cover the branches it happens to miss —
/// deflated metafiles, the two-digest variants, nesting, and hostile input.
/// </remarks>
public sealed class EscherPictureTests
{
    private const int DggContainer = 0xF000;
    private const int BStoreContainer = 0xF001;
    private const int DgContainer = 0xF002;
    private const int SpgrContainer = 0xF003;
    private const int SpContainer = 0xF004;
    private const int Fsp = 0xF00A;
    private const int Fopt = 0xF00B;
    private const int ClientAnchor = 0xF010;
    private const int FBSE = 0xF007;
    private const int BlipPng = 0xF01E;
    private const int BlipWmf = 0xF01B;

    /// <summary>A raster image sits behind one digest and a tag byte, and comes back verbatim.</summary>
    [Fact]
    public void ARasterImageIsHandedBackUntouched()
    {
        byte[] png = [0x89, 0x50, 0x4E, 0x47, 1, 2, 3, 4, 5];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(RasterBlip(png, twoDigests: false)));

        Assert.Equal(png, Assert.Single(store).Content.ToArray());
    }

    /// <summary>
    /// The odd instance values carry a second digest before the data. Reading one as the other
    /// shifts everything by sixteen bytes, which is a corruption no length check would catch.
    /// </summary>
    [Fact]
    public void TheSecondDigestIsSteppedOverWhenItIsThere()
    {
        byte[] png = [0x89, 0x50, 0x4E, 0x47, 9, 8, 7];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(RasterBlip(png, twoDigests: true)));

        Assert.Equal(png, Assert.Single(store).Content.ToArray());
    }

    /// <summary>
    /// A metafile is stored deflated, and inflating it is the one case where the bytes handed
    /// back are not the bytes read. There is no corpus file to check this against.
    /// </summary>
    [Fact]
    public void ADeflatedMetafileIsInflated()
    {
        byte[] wmf = [0xD7, 0xCD, 0xC6, 0x9A, .. Enumerable.Repeat((byte)0x42, 200)];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(MetafileBlip(wmf, compressed: true)));

        Assert.Equal(wmf, Assert.Single(store).Content.ToArray());
    }

    /// <summary>A metafile the writer chose not to compress is read as it lies.</summary>
    [Fact]
    public void AnUncompressedMetafileIsReadAsItLies()
    {
        byte[] wmf = [0xD7, 0xCD, 0xC6, 0x9A, 1, 2, 3];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(MetafileBlip(wmf, compressed: false)));

        Assert.Equal(wmf, Assert.Single(store).Content.ToArray());
    }

    /// <summary>
    /// A picture inside a group, inside the drawing. Shapes sit at whatever depth the file put
    /// them at, so the walk has to descend rather than look at one level.
    /// </summary>
    [Fact]
    public void APictureNestedInGroupsIsStillFound()
    {
        byte[] shape = Container(SpContainer, [.. PictureShape(blip: 1, "B3", "E9")]);
        byte[] drawing = Container(DgContainer, Container(SpgrContainer, Container(SpgrContainer, shape)));

        var shapes = new List<EscherPictures.Shape>();
        EscherPictures.ReadShapes(drawing, shapes);

        Assert.Equal(1, Assert.Single(shapes).BlipIndex);
    }

    /// <summary>
    /// A tree that says it nests forever must not take the stack with it. A workbook is
    /// untrusted input, and the depth of its shape tree is the file's to state.
    /// </summary>
    [Fact]
    public void AbsurdlyDeepNestingIsRefusedRatherThanFollowed()
    {
        // Every container declares a body running to the end of the buffer, so each level opens
        // another that looks exactly the same. Nothing here is a picture; the point is to return.
        var bytes = new List<byte>();
        for (int i = 0; i < 4000; i++)
            bytes.AddRange(Header(0x0F, 0, SpgrContainer, (4000 - i - 1) * 8));

        var shapes = new List<EscherPictures.Shape>();
        EscherPictures.ReadShapes(bytes.ToArray(), shapes);

        Assert.Empty(shapes);
    }

    /// <summary>
    /// A shape naming a BLIP the store does not have produces no picture. Handing back an empty
    /// image, or the wrong one, would be worse than leaving it out.
    /// </summary>
    [Fact]
    public void AShapePointingPastTheStoreIsDropped()
    {
        var sheet = Workbook.Create().Sheets.Add("One");
        EscherPictures.Attach(sheet, [], [new EscherPictures.Shape(3, default, null)]);

        Assert.Empty(sheet.Pictures);
    }

    /// <summary>
    /// The anchor's offsets are fractions of a cell, and turning them into a length needs the
    /// width that cell actually has — so a widened column moves the picture's edge with it.
    /// </summary>
    [Fact]
    public void TheAnchorOffsetFollowsTheColumnItIsMeasuredAgainst()
    {
        byte[] png = [0x89, 0x50, 0x4E, 0x47, 1];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(RasterBlip(png, twoDigests: false)));

        var narrow = Workbook.Create().Sheets.Add("Narrow");
        var wide = Workbook.Create().Sheets.Add("Wide");
        wide.SetColumnWidth(2, 40);

        // Half way into column B, which is column index 1.
        var anchor = new EscherPictures.SheetAnchor(0x0003, 1, 512, 2, 0, 4, 0, 8, 0);
        EscherPictures.Attach(narrow, store, [new EscherPictures.Shape(1, anchor, null)]);
        EscherPictures.Attach(wide, store, [new EscherPictures.Shape(1, anchor, null)]);

        Assert.True(
            wide.Pictures[0].Anchor.FromOffset.X > narrow.Pictures[0].Anchor.FromOffset.X,
            "A wider column should put the same fractional offset further right.");
    }

    /// <summary>The two flag bits together are what tells the three move behaviours apart.</summary>
    [Theory]
    [InlineData(0x0003, PictureMoveBehavior.MoveAndResize)]
    [InlineData(0x0001, PictureMoveBehavior.MoveOnly)]
    [InlineData(0x0000, PictureMoveBehavior.Fixed)]
    public void TheAnchorFlagsDecideWhatHappensWhenCellsMove(int flags, PictureMoveBehavior expected)
    {
        byte[] png = [0x89, 0x50, 0x4E, 0x47, 1];
        List<EscherPictures.Blip> store = EscherPictures.ReadStore(DrawingGroup(RasterBlip(png, twoDigests: false)));
        var sheet = Workbook.Create().Sheets.Add("Flags");

        EscherPictures.Attach(
            sheet, store, [new EscherPictures.Shape(1, new EscherPictures.SheetAnchor((ushort)flags, 0, 0, 0, 0, 2, 0, 3, 0), null)]);

        Assert.Equal(expected, sheet.Pictures[0].Anchor.MoveBehavior);
    }

    /// <summary>Wraps a BLIP in the store containers a drawing group holds it in.</summary>
    private static byte[] DrawingGroup(byte[] blip) =>
        Container(DggContainer, Container(BStoreContainer, StoreEntry(blip)));

    /// <summary>An <c>OfficeArtFBSE</c> with the BLIP embedded in it (MS-ODRAW 2.2.32).</summary>
    private static byte[] StoreEntry(byte[] blip)
    {
        var body = new List<byte>();
        body.Add(0x06);                                 // btWin32.
        body.Add(0x06);                                 // btMacOS.
        body.AddRange(new byte[16]);                    // rgbUid.
        body.AddRange([0xFF, 0x00]);                    // tag.
        body.AddRange(BitConverter.GetBytes(blip.Length));
        body.AddRange(BitConverter.GetBytes(1));        // cRef.
        body.AddRange(BitConverter.GetBytes(0));        // foDelay.
        body.AddRange([0, 0, 0, 0]);                    // unused1, cbName, unused2, unused3.
        body.AddRange(blip);

        return [.. Header(0x02, 0x06E0, FBSE, body.Count), .. body];
    }

    /// <summary>An <c>OfficeArtBlipPNG</c>: digests, a tag byte, then the file (MS-ODRAW 2.2.28).</summary>
    private static byte[] RasterBlip(byte[] content, bool twoDigests)
    {
        var body = new List<byte>();
        body.AddRange(new byte[twoDigests ? 32 : 16]);
        body.Add(0xFF);
        body.AddRange(content);

        return [.. Header(0x00, twoDigests ? 0x06E1 : 0x06E0, BlipPng, body.Count), .. body];
    }

    /// <summary>An <c>OfficeArtBlipWMF</c>: a digest, a 34-byte header, then the deflated file.</summary>
    private static byte[] MetafileBlip(byte[] content, bool compressed)
    {
        byte[] stored = compressed ? Deflate(content) : content;

        var body = new List<byte>();
        body.AddRange(new byte[16]);                    // rgbUid1.
        body.AddRange(BitConverter.GetBytes(content.Length));
        body.AddRange(new byte[16]);                    // rcBounds.
        body.AddRange(new byte[8]);                     // ptSize.
        body.AddRange(BitConverter.GetBytes(stored.Length));
        body.Add(compressed ? (byte)0x00 : (byte)0xFE); // compression.
        body.Add(0xFE);                                 // filter.
        body.AddRange(stored);

        return [.. Header(0x00, 0x0216, BlipWmf, body.Count), .. body];
    }

    private static byte[] Deflate(byte[] content)
    {
        using var output = new MemoryStream();
        using (var compressor = new ZLibStream(output, CompressionLevel.Optimal, leaveOpen: true))
            compressor.Write(content);
        return output.ToArray();
    }

    /// <summary>A picture shape: its type, the property naming its BLIP, and where it sits.</summary>
    private static byte[] PictureShape(int blip, string from, string to)
    {
        CellRef start = CellRef.Parse(from);
        CellRef end = CellRef.Parse(to);

        var properties = new List<byte>();
        properties.AddRange([0x04, 0x41]);              // pib, with fBid set.
        properties.AddRange(BitConverter.GetBytes(blip));

        var anchor = new List<byte>();
        anchor.AddRange(BitConverter.GetBytes((ushort)0x0003));
        foreach (int value in new[] { start.Column - 1, 0, start.Row - 1, 0, end.Column - 1, 0, end.Row - 1, 0 })
            anchor.AddRange(BitConverter.GetBytes((ushort)value));

        return
        [
            .. Header(0x02, 75, Fsp, 8), .. new byte[8],
            .. Header(0x03, 1, Fopt, properties.Count), .. properties,
            .. Header(0x00, 0, ClientAnchor, anchor.Count - 2), .. anchor[..^2],
        ];
    }

    private static byte[] Container(int type, byte[] body) => [.. Header(0x0F, 0, type, body.Length), .. body];

    /// <summary>An <c>OfficeArtRecordHeader</c>: version and instance packed, type, then length.</summary>
    private static byte[] Header(int version, int instance, int type, int length)
    {
        var header = new byte[8];
        BinaryPrimitives.WriteUInt16LittleEndian(header, (ushort)((instance << 4) | version));
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(2), (ushort)type);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(4), (uint)length);
        return header;
    }
}
