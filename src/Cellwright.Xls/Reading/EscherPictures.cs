using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Pulls the images out of the Office Art layer of a legacy workbook (MS-ODRAW).
/// </summary>
/// <remarks>
/// The bytes of every image in the file sit together in the workbook's own drawing group, in a
/// BLIP store; a sheet's shapes then name one by its position in that store. So this is read in
/// two passes, and they are in different substreams: the store comes from the globals, the
/// shapes from each sheet.
/// <para>
/// Nothing here decodes an image. A raster BLIP wraps the file exactly as it was imported and
/// the wrapper is simply stepped over; a metafile is the one exception, because the format
/// deflates it, and inflating is still handing back the original bytes rather than re-encoding
/// them.
/// </para>
/// </remarks>
internal static class EscherPictures
{
    private const int BStoreContainer = 0xF001;
    private const int FBSE = 0xF007;
    private const int SpContainer = 0xF004;
    private const int Fsp = 0xF00A;
    private const int Fopt = 0xF00B;
    private const int ClientAnchor = 0xF010;

    /// <summary>Property that names which BLIP a picture shape shows (MS-ODRAW 2.3.7.2).</summary>
    private const int PibProperty = 0x0104;

    /// <summary>Property holding the shape's name (MS-ODRAW 2.3.4.20, wzName).</summary>
    private const int NameProperty = 0x0380;

    /// <summary>
    /// Reads the BLIP store out of a <c>MsoDrawingGroup</c> record. The result is indexed the
    /// way a shape names it: one-based, in store order.
    /// </summary>
    public static List<Blip> ReadStore(ReadOnlyMemory<byte> body)
    {
        var blips = new List<Blip>();
        if (Find(body, BStoreContainer) is not { } store)
            return blips;

        foreach (Record entry in Children(body, store.Start, store.Length))
        {
            if (entry.Type == FBSE)
                blips.Add(ReadEntry(body, entry));
        }

        return blips;
    }

    /// <summary>
    /// Reads the picture shapes of one sheet's drawing. Shapes that are not pictures — the text
    /// boxes behind notes, controls, charts — are left to their own readers.
    /// </summary>
    public static void ReadShapes(ReadOnlyMemory<byte> body, List<Shape> shapes) =>
        Walk(body, 0, body.Length, shapes);

    /// <summary>
    /// Joins the record bodies an Office Art tree arrived in.
    /// </summary>
    /// <remarks>
    /// The drawing layer does not fit the record format it travels in, and neither side gives
    /// way: the tree is chopped at whatever byte the record runs out at, so a container's
    /// declared length routinely runs past the end of the record carrying it. A workbook's
    /// image store continues into a second <c>MsoDrawingGroup</c> and a sheet's shapes continue
    /// into further <c>MsoDrawing</c> records, and in both cases parsing a piece on its own
    /// reads only as far as that piece happens to reach.
    /// </remarks>
    public static byte[] Join(List<byte[]> pieces)
    {
        if (pieces.Count == 1)
            return pieces[0];

        var joined = new byte[pieces.Sum(static piece => piece.Length)];
        int offset = 0;
        foreach (byte[] piece in pieces)
        {
            piece.CopyTo(joined, offset);
            offset += piece.Length;
        }

        return joined;
    }

    /// <summary>
    /// Turns what the two passes found into pictures on the sheet, converting each anchor from
    /// the fractions of a cell that BIFF measures in.
    /// </summary>
    public static void Attach(Worksheet sheet, IReadOnlyList<Blip> store, IReadOnlyList<Shape> shapes)
    {
        foreach (Shape shape in shapes)
        {
            if (shape.BlipIndex < 1 || shape.BlipIndex > store.Count)
                continue;

            Blip blip = store[shape.BlipIndex - 1];
            if (blip.Content.Length == 0)
                continue;

            // The public Add is deliberate, as it is for notes: a legacy file has no drawing
            // part to replay on save, so the model has to own these pictures for the xlsx
            // writer to have anything to write.
            sheet.Pictures.Add(new Picture
            {
                Content = blip.Content,
                Format = PictureCollection.DetectFormat(blip.Content.Span),
                Anchor = shape.Anchor.ToAnchor(sheet),
                Name = shape.Name,
            });
        }
    }

    /// <summary>One image from the store: the bytes, ready to hand to the model.</summary>
    internal readonly record struct Blip(ReadOnlyMemory<byte> Content);

    /// <summary>A picture shape found on a sheet, before its BLIP has been looked up.</summary>
    internal readonly record struct Shape(int BlipIndex, SheetAnchor Anchor, string? Name);

    /// <summary>
    /// An <c>OfficeArtClientAnchorSheet</c> (MS-XLS 2.5.193) as the file states it: cells, and
    /// offsets into them measured as fractions — 1024ths of a column's width, 256ths of a row's
    /// height. Nothing here is in EMU yet, because a fraction of a cell cannot be turned into a
    /// length without knowing how wide that cell is.
    /// </summary>
    internal readonly record struct SheetAnchor(
        ushort Flags, int ColumnLeft, int OffsetLeft, int RowTop, int OffsetTop,
        int ColumnRight, int OffsetRight, int RowBottom, int OffsetBottom)
    {
        public static SheetAnchor Read(ReadOnlySpan<byte> body) => body.Length < 18
            ? default
            : new SheetAnchor(
                BinaryPrimitives.ReadUInt16LittleEndian(body),
                BinaryPrimitives.ReadUInt16LittleEndian(body[2..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[4..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[6..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[8..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[10..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[12..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[14..]),
                BinaryPrimitives.ReadUInt16LittleEndian(body[16..]));

        /// <summary>
        /// Converts to the anchor the model holds. The cells carry over exactly; the offsets
        /// have to be multiplied out against the widths and heights actually in effect, which
        /// is the one place a legacy picture's position is reconstructed rather than read.
        /// </summary>
        public PictureAnchor ToAnchor(Worksheet sheet) => new()
        {
            Kind = PictureAnchorKind.TwoCell,
            From = Cell(RowTop, ColumnLeft),
            FromOffset = (ColumnEmu(sheet, ColumnLeft, OffsetLeft), RowEmu(sheet, RowTop, OffsetTop)),
            To = Cell(RowBottom, ColumnRight),
            ToOffset = (ColumnEmu(sheet, ColumnRight, OffsetRight), RowEmu(sheet, RowBottom, OffsetBottom)),

            // fMove alone is "move but don't size"; both bits is the ordinary "move and size".
            MoveBehavior = (Flags & 0x0001) == 0
                ? PictureMoveBehavior.Fixed
                : (Flags & 0x0002) != 0 ? PictureMoveBehavior.MoveAndResize : PictureMoveBehavior.MoveOnly,
        };

        private static CellRef Cell(int row, int column) => new(
            Math.Clamp(row + 1, 1, SheetLimits.MaxRow),
            Math.Clamp(column + 1, 1, SheetLimits.MaxColumn));

        private static long ColumnEmu(Worksheet sheet, int column, int offset)
        {
            double characters = sheet.ColumnTable.GetValueOrDefault(column + 1).Width ?? DefaultColumnWidth;
            double pixels = Math.Round((characters * MaximumDigitWidth) + ColumnPadding);
            return (long)Math.Round(pixels * Emu.PerPixel * offset / 1024.0);
        }

        private static long RowEmu(Worksheet sheet, int row, int offset)
        {
            double points = sheet.RowTable.GetValueOrDefault(row + 1).Height ?? DefaultRowHeight;
            return (long)Math.Round(points * Emu.PerPoint * offset / 256.0);
        }

        /// <summary>Width of a column Excel calls "8.43 characters", in the Calibri 11 default.</summary>
        private const double DefaultColumnWidth = 8.43;

        /// <summary>Height of a row Excel calls default, in points.</summary>
        private const double DefaultRowHeight = 15.0;

        /// <summary>Pixels the widest digit of the default font occupies, which a width is counted in.</summary>
        private const double MaximumDigitWidth = 7.0;

        /// <summary>Pixels of padding a column carries beside its characters.</summary>
        private const double ColumnPadding = 5.0;
    }

    /// <summary>
    /// Walks a shape tree, collecting the pictures.
    /// </summary>
    /// <remarks>
    /// Groups nest, and the obvious way to follow them is recursion — but the depth here comes
    /// from the file, and a file that says a group contains a group all the way down would take
    /// the stack with it. A workbook is untrusted input, so the nesting is followed on a stack
    /// of our own with a ceiling on it.
    /// </remarks>
    private static void Walk(ReadOnlyMemory<byte> owner, int start, int length, List<Shape> shapes)
    {
        const int MaxDepth = 64;
        var pending = new Stack<(int Start, int Length, int Depth)>();
        pending.Push((start, length, 0));

        while (pending.TryPop(out (int Start, int Length, int Depth) frame))
        {
            foreach (Record record in Children(owner, frame.Start, frame.Length))
            {
                if (record.Type == SpContainer)
                {
                    if (ReadShape(owner, record) is { } shape)
                        shapes.Add(shape);
                }
                else if (record.IsContainer && frame.Depth < MaxDepth)
                {
                    // The drawing container, the group containers under it, and anything else
                    // that holds records: a shape can be at any of these depths.
                    pending.Push((record.Start, record.Length, frame.Depth + 1));
                }
            }
        }
    }

    /// <summary>
    /// Reads one shape container, keeping it only if it displays an image.
    /// </summary>
    /// <remarks>
    /// What makes a shape a picture is the <c>pib</c> property, not its shape type. A plain
    /// image is a picture frame, but an embedded object or a form control shows its face the
    /// same way — through a BLIP index — and to someone looking at the sheet that is a picture
    /// in exactly the same sense.
    /// </remarks>
    private static Shape? ReadShape(ReadOnlyMemory<byte> owner, in Record container)
    {
        int blip = 0;
        string? name = null;
        SheetAnchor anchor = default;

        foreach (Record child in Children(owner, container.Start, container.Length))
        {
            switch (child.Type)
            {
                case Fopt:
                    (blip, name) = ReadProperties(child.Span(owner), child.Instance);
                    break;
                case ClientAnchor:
                    anchor = SheetAnchor.Read(child.Span(owner));
                    break;
                default:
                    break;
            }
        }

        return blip > 0 ? new Shape(blip, anchor, name) : null;
    }

    /// <summary>
    /// Reads what a property table says about the picture: which BLIP it shows, and what the
    /// shape is called.
    /// </summary>
    /// <remarks>
    /// Every entry is a fixed six bytes, and the complex ones keep their payload after the whole
    /// table rather than inline — in table order, each as long as its own entry declared. So the
    /// name can only be found by adding up the lengths of the complex properties before it.
    /// </remarks>
    private static (int Blip, string? Name) ReadProperties(ReadOnlySpan<byte> body, int count)
    {
        int blip = 0;
        int payload = count * 6;
        int nameStart = -1;
        int nameLength = 0;

        for (int i = 0; i < count && ((i + 1) * 6) <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(i * 6, 6);
            ushort header = BinaryPrimitives.ReadUInt16LittleEndian(entry);
            uint value = BinaryPrimitives.ReadUInt32LittleEndian(entry[2..]);
            int id = header & 0x3FFF;
            bool complex = (header & 0x8000) != 0;

            if (!complex && id == PibProperty)
            {
                blip = (int)value;
                continue;
            }

            if (!complex)
                continue;

            if (id == NameProperty)
            {
                nameStart = payload;
                nameLength = (int)value;
            }

            payload += (int)value;
        }

        return (blip, ReadName(body, nameStart, nameLength));
    }

    /// <summary>The shape name, a NUL-terminated UTF-16 string sitting after the property table.</summary>
    private static string? ReadName(ReadOnlySpan<byte> body, int start, int length)
    {
        if (start < 0 || length < 2 || start + length > body.Length)
            return null;

        string text = Encoding.Unicode.GetString(body.Slice(start, length & ~1)).TrimEnd('\0');
        return text.Length == 0 ? null : text;
    }

    /// <summary>
    /// Reads one store entry. The image may be embedded in the entry or held in a stream of its
    /// own; only the embedded form is read, because the other lives outside this record.
    /// </summary>
    private static Blip ReadEntry(ReadOnlyMemory<byte> owner, in Record entry)
    {
        // btWin32, btMacOS, the digest, tag, size, cRef, foDelay, then four bytes of which one
        // is the name length. The name, when there is one, sits between them and the BLIP.
        if (entry.Length < 36)
            return default;

        int nameLength = entry.Span(owner)[33];
        int offset = 36 + nameLength;
        if (offset >= entry.Length)
            return default;

        foreach (Record blip in Children(owner, entry.Start + offset, entry.Length - offset))
            return ReadBlip(owner, blip);

        return default;
    }

    /// <summary>
    /// Unwraps a BLIP. Every kind starts with one digest, and the odd instance values carry a
    /// second; past those, a raster format has a single tag byte and then the file itself,
    /// while a metafile has a header saying how the file was deflated.
    /// </summary>
    private static Blip ReadBlip(ReadOnlyMemory<byte> owner, in Record record)
    {
        int offset = (record.Instance & 1) != 0 ? 32 : 16;
        if (offset >= record.Length)
            return default;

        switch (record.Type)
        {
            // PNG, JPEG, DIB, TIFF and the CMYK JPEG: the file follows the tag byte verbatim.
            case 0xF01D or 0xF01E or 0xF01F or 0xF029 or 0xF02A:
                return new Blip(record.Slice(owner, offset + 1));

            // EMF, WMF and PICT, whose file the header before it says how to inflate.
            case 0xF01A or 0xF01B or 0xF01C:
                return ReadMetafile(owner, record, offset);

            default:
                return default;
        }
    }

    private static Blip ReadMetafile(ReadOnlyMemory<byte> owner, in Record record, int offset)
    {
        const int HeaderSize = 34;
        if (offset + HeaderSize > record.Length)
            return default;

        ReadOnlySpan<byte> header = record.Span(owner)[offset..];
        byte compression = header[32];
        ReadOnlyMemory<byte> stored = record.Slice(owner, offset + HeaderSize);
        if (compression == 0xFE)
            return new Blip(stored);

        try
        {
            using var input = new MemoryStream(stored.ToArray(), writable: false);
            using var inflate = new ZLibStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream(1024);
            inflate.CopyTo(output);
            return new Blip(output.ToArray());
        }
        catch (InvalidDataException)
        {
            // A metafile that will not inflate is one this reader cannot hand back honestly.
            return default;
        }
    }

    /// <summary>
    /// One Office Art record: its header fields, and where its body lies in the buffer the
    /// whole tree was read from. It holds a range rather than the bytes so that walking a tree
    /// copies nothing.
    /// </summary>
    private readonly record struct Record(int Version, int Instance, int Type, int Start, int Length)
    {
        /// <summary>Whether the record holds other records rather than data of its own.</summary>
        public bool IsContainer => Version == 0x0F;

        public ReadOnlySpan<byte> Span(ReadOnlyMemory<byte> owner) => owner.Span.Slice(Start, Length);

        /// <summary>The body from an offset on, as memory the model can hold onto.</summary>
        public ReadOnlyMemory<byte> Slice(ReadOnlyMemory<byte> owner, int offset) =>
            offset >= Length ? default : owner.Slice(Start + offset, Length - offset);
    }

    /// <summary>
    /// Walks the records directly inside a range. Every Office Art record opens with the same
    /// eight bytes, which is what makes a tree of them walkable without knowing any of them.
    /// </summary>
    private static List<Record> Children(ReadOnlyMemory<byte> owner, int start, int length)
    {
        var records = new List<Record>();
        int offset = start;
        int end = start + length;
        while (offset + 8 <= end)
        {
            ReadOnlySpan<byte> header = owner.Span.Slice(offset, 8);
            int versionAndInstance = BinaryPrimitives.ReadUInt16LittleEndian(header);
            int type = BinaryPrimitives.ReadUInt16LittleEndian(header[2..]);
            uint declared = BinaryPrimitives.ReadUInt32LittleEndian(header[4..]);

            // A container's length may overrun its parent in a file the writer chopped into
            // records; clamping keeps the walk inside the buffer instead of giving up on it.
            int body = (int)Math.Min(declared, (uint)(end - offset - 8));
            records.Add(new Record(versionAndInstance & 0x0F, versionAndInstance >> 4, type, offset + 8, body));
            offset += 8 + body;
        }

        return records;
    }

    /// <summary>
    /// Finds a container anywhere in the tree, descending into the containers on the way.
    /// </summary>
    /// <remarks>
    /// The BLIP store is not a top-level record: it sits inside the drawing group container,
    /// beside the shape-identifier block and the colour tables, and not at a fixed position
    /// among them. Descending is what finds it — and what a record says about itself is enough
    /// to know whether descending is even possible, since a container declares version 0xF.
    /// </remarks>
    private static Record? Find(ReadOnlyMemory<byte> owner, int type)
    {
        const int MaxDepth = 16;
        var pending = new Queue<(int Start, int Length, int Depth)>();
        pending.Enqueue((0, owner.Length, 0));

        while (pending.TryDequeue(out (int Start, int Length, int Depth) frame))
        {
            foreach (Record record in Children(owner, frame.Start, frame.Length))
            {
                if (record.Type == type)
                    return record;
                if (record.IsContainer && frame.Depth < MaxDepth)
                    pending.Enqueue((record.Start, record.Length, frame.Depth + 1));
            }
        }

        return null;
    }
}
