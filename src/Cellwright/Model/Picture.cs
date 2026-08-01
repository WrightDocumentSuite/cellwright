using System.Collections;

namespace Cellwright;

/// <summary>
/// How a picture is pinned to the sheet.
/// </summary>
public enum PictureAnchorKind
{
    /// <summary>Corners are pinned to two cells, so the picture resizes with them.</summary>
    TwoCell,

    /// <summary>The top-left corner is pinned to a cell and the size is fixed.</summary>
    OneCell,

    /// <summary>Both position and size are fixed, independent of the grid.</summary>
    Absolute,
}

/// <summary>
/// Where a picture sits on the sheet, in the terms the file uses.
/// </summary>
/// <remarks>
/// Offsets and sizes are in English Metric Units, the unit every drawing part is written in:
/// 914 400 to the inch, 12 700 to the point. Converting them to something friendlier would mean
/// guessing at the row heights and column widths in effect, so the model keeps what the file
/// said and offers <see cref="Emu"/> to do the arithmetic.
/// </remarks>
public sealed class PictureAnchor
{
    /// <summary>Which of the three ways the picture is pinned.</summary>
    public required PictureAnchorKind Kind { get; init; }

    /// <summary>Cell holding the top-left corner. Unused by <see cref="PictureAnchorKind.Absolute"/>.</summary>
    public CellRef From { get; set; }

    /// <summary>Offset into <see cref="From"/>, in EMU, right and down from its top-left corner.</summary>
    public (long X, long Y) FromOffset { get; set; }

    /// <summary>Cell holding the bottom-right corner. Only for <see cref="PictureAnchorKind.TwoCell"/>.</summary>
    public CellRef To { get; set; }

    /// <summary>Offset into <see cref="To"/>, in EMU.</summary>
    public (long X, long Y) ToOffset { get; set; }

    /// <summary>
    /// Size in EMU, for the two kinds that carry one. A two-cell anchor has no size of its own:
    /// its size is whatever the cells between its corners come to.
    /// </summary>
    public (long Width, long Height) Extent { get; set; }

    /// <summary>Position in EMU from the top-left of the sheet, for an absolute anchor.</summary>
    public (long X, long Y) Position { get; set; }

    /// <summary>What happens to the picture when the cells under it move or resize.</summary>
    public PictureMoveBehavior MoveBehavior { get; set; } = PictureMoveBehavior.MoveAndResize;

    /// <summary>Pins a picture between two cells, so it moves and resizes with them.</summary>
    public static PictureAnchor Between(CellRef from, CellRef to) =>
        new() { Kind = PictureAnchorKind.TwoCell, From = from, To = to };

    /// <summary>Pins the top-left corner to a cell and gives the picture a fixed size.</summary>
    public static PictureAnchor At(CellRef cell, long width, long height) =>
        new() { Kind = PictureAnchorKind.OneCell, From = cell, Extent = (width, height) };

    /// <inheritdoc/>
    public override string ToString() => Kind switch
    {
        PictureAnchorKind.TwoCell => $"{From}:{To}",
        PictureAnchorKind.OneCell => $"{From} ({Extent.Width}x{Extent.Height} EMU)",
        _ => $"({Position.X},{Position.Y}) {Extent.Width}x{Extent.Height} EMU",
    };
}

/// <summary>What a picture does when the cells underneath it change.</summary>
public enum PictureMoveBehavior
{
    /// <summary>Moves and resizes with the cells. The default for a two-cell anchor.</summary>
    MoveAndResize,

    /// <summary>Moves with the cells but keeps its size.</summary>
    MoveOnly,

    /// <summary>Stays put whatever the cells do.</summary>
    Fixed,
}

/// <summary>Image encodings a spreadsheet can carry.</summary>
public enum PictureFormat
{
    /// <summary>Not one of the encodings this library recognizes.</summary>
    Unknown,

    /// <summary>Portable Network Graphics.</summary>
    Png,

    /// <summary>JPEG.</summary>
    Jpeg,

    /// <summary>Graphics Interchange Format.</summary>
    Gif,

    /// <summary>Windows bitmap.</summary>
    Bmp,

    /// <summary>Tagged Image File Format.</summary>
    Tiff,

    /// <summary>Windows Enhanced Metafile.</summary>
    Emf,

    /// <summary>Windows Metafile.</summary>
    Wmf,
}

/// <summary>
/// An image placed on a worksheet.
/// </summary>
/// <remarks>
/// The bytes are handed over exactly as the file stored them, and written back the same way.
/// This library does not decode, re-encode or resample an image: a picture that arrived as a
/// progressive JPEG leaves as the same progressive JPEG, byte for byte.
/// </remarks>
public sealed class Picture
{
    /// <summary>The encoded image, exactly as stored.</summary>
    public required ReadOnlyMemory<byte> Content { get; init; }

    /// <summary>Encoding of <see cref="Content"/>.</summary>
    public required PictureFormat Format { get; init; }

    /// <summary>Where the picture sits.</summary>
    public required PictureAnchor Anchor { get; init; }

    /// <summary>Name shown in the selection pane, e.g. <c>Picture 1</c>.</summary>
    public string? Name { get; init; }

    /// <summary>Alternative text for accessibility, or <see langword="null"/>.</summary>
    public string? Description { get; init; }

    /// <summary>Whether the picture is hidden.</summary>
    public bool IsHidden { get; init; }

    /// <inheritdoc/>
    public override string ToString() => $"{Name ?? Format.ToString()} at {Anchor}";
}

/// <summary>
/// The pictures of one worksheet, in the order the drawing part lists them — which is the order
/// they are stacked, back to front.
/// </summary>
public sealed class PictureCollection : IReadOnlyList<Picture>
{
    private readonly List<Picture> _pictures = [];

    /// <inheritdoc/>
    public int Count => _pictures.Count;

    /// <inheritdoc/>
    public Picture this[int index] => _pictures[index];

    /// <summary>
    /// Whether the model, rather than the loaded file, owns the sheet's pictures. Until
    /// something changes, the original drawing part travels through untouched, which keeps the
    /// shapes and charts beside the pictures that the model has no room for.
    /// </summary>
    internal bool IsModified { get; private set; }

    /// <summary>Places a picture on the sheet.</summary>
    public Picture Add(Picture picture)
    {
        ArgumentNullException.ThrowIfNull(picture);
        if (picture.Content.IsEmpty)
            throw new ArgumentException("A picture has no content.", nameof(picture));

        _pictures.Add(picture);
        IsModified = true;
        return picture;
    }

    /// <summary>
    /// Places an image between two cells, so it moves and resizes with them. The encoding is
    /// taken from the bytes themselves rather than from a file name, which may lie.
    /// </summary>
    public Picture Add(ReadOnlyMemory<byte> content, CellRef from, CellRef to, string? name = null) =>
        Add(new Picture
        {
            Content = content,
            Format = DetectFormat(content.Span),
            Anchor = PictureAnchor.Between(from, to),
            Name = name,
        });

    /// <summary>Places an image between two cells given as A1 references.</summary>
    public Picture Add(ReadOnlyMemory<byte> content, string from, string to, string? name = null) =>
        Add(content, CellRef.Parse(from), CellRef.Parse(to), name);

    /// <summary>Removes a picture. Returns whether it was there.</summary>
    public bool Remove(Picture picture)
    {
        bool removed = _pictures.Remove(picture);
        IsModified |= removed;
        return removed;
    }

    /// <summary>Removes every picture.</summary>
    public void Clear()
    {
        if (_pictures.Count == 0)
            return;

        _pictures.Clear();
        IsModified = true;
    }

    /// <summary>Adds a picture read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(Picture picture) => _pictures.Add(picture);

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit moved a picture.</summary>
    internal void MarkModified() => IsModified = true;

    /// <summary>
    /// Moves the pictures whose cells the mapping moves, dropping the ones it maps to nothing.
    /// Used when rows or columns are inserted or deleted underneath them.
    /// </summary>
    /// <remarks>
    /// The collection is only marked as edited when a picture actually moves, and that matters
    /// more here than elsewhere: an edited collection means the drawing part is written afresh,
    /// which costs the shapes and charts that shared it. A sheet whose pictures all sit above
    /// an inserted row keeps its original part.
    /// </remarks>
    internal void Reanchor(Func<CellRef, CellRef?> map)
    {
        bool moved = false;
        for (int i = _pictures.Count - 1; i >= 0; i--)
        {
            PictureAnchor anchor = _pictures[i].Anchor;
            if (anchor.Kind == PictureAnchorKind.Absolute || anchor.MoveBehavior == PictureMoveBehavior.Fixed)
                continue;

            if (map(anchor.From) is not { } from)
            {
                _pictures.RemoveAt(i);
                moved = true;
                continue;
            }

            CellRef to = anchor.To;
            if (anchor.Kind == PictureAnchorKind.TwoCell && map(to) is { } mappedTo)
                to = mappedTo;

            if (from == anchor.From && to == anchor.To)
                continue;

            anchor.From = from;
            anchor.To = to;
            moved = true;
        }

        IsModified |= moved;
    }

    /// <inheritdoc/>
    public IEnumerator<Picture> GetEnumerator() => _pictures.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Recognizes an image by its leading bytes. Every format here announces itself in the
    /// first few bytes, which is both what Excel relies on and what makes the extension of the
    /// part a formality.
    /// </summary>
    internal static PictureFormat DetectFormat(ReadOnlySpan<byte> content)
    {
        // Spelled out rather than written as a u8 literal: the first byte of the PNG signature
        // is 0x89, which as a string literal is the character U+0089 and encodes to two bytes.
        if (content.Length >= 8 && content[0] == 0x89 && content[1..8].SequenceEqual("PNG\r\n\x1a\n"u8))
            return PictureFormat.Png;
        if (content.Length >= 3 && content[0] == 0xFF && content[1] == 0xD8 && content[2] == 0xFF)
            return PictureFormat.Jpeg;
        if (content.Length >= 6 && (content[..6].SequenceEqual("GIF87a"u8) || content[..6].SequenceEqual("GIF89a"u8)))
            return PictureFormat.Gif;
        if (content.Length >= 2 && content[0] == 0x42 && content[1] == 0x4D)
            return PictureFormat.Bmp;
        if (content.Length >= 4 && (content[..4].SequenceEqual("II\x2a\0"u8) || content[..4].SequenceEqual("MM\0\x2a"u8)))
            return PictureFormat.Tiff;
        if (content.Length >= 44 && content[40] == 0x20 && content[41] == 0x45 && content[42] == 0x4D && content[43] == 0x46)
            return PictureFormat.Emf;
        if (content.Length >= 4 && content[0] == 0xD7 && content[1] == 0xCD && content[2] == 0xC6 && content[3] == 0x9A)
            return PictureFormat.Wmf;

        // A metafile without the placeable header starts with its own record header instead.
        if (content.Length >= 4 && content[0] == 0x01 && content[1] == 0x00 && content[2] == 0x09 && content[3] == 0x00)
            return PictureFormat.Wmf;

        return PictureFormat.Unknown;
    }
}

/// <summary>
/// English Metric Units, the unit drawing parts measure in. There are 914 400 to the inch,
/// which makes them divide evenly into both inches and centimetres.
/// </summary>
public static class Emu
{
    /// <summary>EMU in one inch.</summary>
    public const long PerInch = 914_400;

    /// <summary>EMU in one point, of which there are 72 to the inch.</summary>
    public const long PerPoint = 12_700;

    /// <summary>EMU in one centimetre.</summary>
    public const long PerCentimetre = 360_000;

    /// <summary>EMU in one pixel at the 96 DPI Excel assumes.</summary>
    public const long PerPixel = 9_525;

    /// <summary>Converts pixels at 96 DPI to EMU.</summary>
    public static long FromPixels(double pixels) => (long)Math.Round(pixels * PerPixel);

    /// <summary>Converts EMU to pixels at 96 DPI.</summary>
    public static double ToPixels(long emu) => (double)emu / PerPixel;

    /// <summary>Converts points to EMU.</summary>
    public static long FromPoints(double points) => (long)Math.Round(points * PerPoint);

    /// <summary>Converts EMU to points.</summary>
    public static double ToPoints(long emu) => (double)emu / PerPoint;
}
