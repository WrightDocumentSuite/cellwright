using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the drawing layer of one sheet: the notes and the pictures, as shapes of a single
/// Office Art drawing.
/// </summary>
/// <remarks>
/// <para>
/// The drawing layer is Office Art, a tree of length-prefixed containers that the format then
/// chops into <c>MsoDrawing</c> records at shape boundaries — so a container's declared length
/// runs past the end of the record carrying it. What this writer produces is the smallest tree
/// Excel accepts: one drawing per sheet, a group shape, and a shape per note or picture.
/// </para>
/// <para>
/// One writer for both kinds, because they share what cannot be shared across writers: the
/// drawing (a sheet has one), the group shape (a drawing has one), and the shape identifiers,
/// which are handed out in one block of 1024 per drawing. Two writers would each claim the
/// block and Excel would repair the file.
/// </para>
/// </remarks>
internal static class EscherDrawingWriter
{
    /// <summary>Office Art record types, in the order they appear in the tree.</summary>
    private const ushort DggContainer = 0xF000;
    private const ushort BStoreContainer = 0xF001;
    private const ushort Dgg = 0xF006;
    private const ushort Bse = 0xF007;
    private const ushort DgContainer = 0xF002;
    private const ushort Dg = 0xF008;
    private const ushort SpgrContainer = 0xF003;
    private const ushort SpContainer = 0xF004;
    private const ushort Spgr = 0xF009;
    private const ushort Sp = 0xF00A;
    private const ushort Opt = 0xF00B;
    private const ushort ClientAnchor = 0xF010;
    private const ushort ClientData = 0xF011;
    private const ushort ClientTextbox = 0xF00D;
    private const ushort SplitMenuColors = 0xF11E;

    /// <summary>Shape type of a text box, which is what a note is drawn as.</summary>
    private const int TextBoxShapeType = 202;

    /// <summary>Shape type of a picture frame.</summary>
    private const int PictureShapeType = 75;

    /// <summary>
    /// The <c>pib</c> property: which store entry the frame shows. The high bit set on the id is
    /// <c>fBid</c>, saying the value is a store index rather than a length — readers mask it off,
    /// but Excel writes it and files are compared against Excel's.
    /// </summary>
    private const int PibProperty = 0x4104;

    /// <summary>Shape ids are handed out in blocks of this size, one block per drawing.</summary>
    private const int ShapeIdsPerDrawing = 1024;

    /// <summary>Bytes one note's shape occupies in the tree, including its text box marker.</summary>
    private const int NoteShapeSize = 134;

    /// <summary>Bytes one picture's shape occupies: the frame, its one property, and its anchor.</summary>
    private const int PictureShapeSize = 72;

    /// <summary>Bytes the group shape that owns every other shape occupies.</summary>
    private const int GroupShapeSize = 48;

    /// <summary>
    /// Writes the drawing of one sheet — pictures first, then notes. Returns how many shapes it
    /// ended up with, zero when the sheet has nothing to draw.
    /// </summary>
    /// <param name="writer">Sheet substream being written.</param>
    /// <param name="sheet">Sheet whose notes and pictures are written.</param>
    /// <param name="drawingId">One-based id of this sheet's drawing.</param>
    /// <param name="context">The write context, which owns the workbook-wide image store.</param>
    public static int Write(BiffWriter writer, Worksheet sheet, int drawingId, XlsWriteContext context)
    {
        List<CellComment> notes = [.. sheet.Comments.OrderBy(static c => c.Cell.Row).ThenBy(static c => c.Cell.Column)];
        List<(Picture Picture, int Blip)> pictures = GatherPictures(sheet, context);
        if (notes.Count == 0 && pictures.Count == 0)
            return 0;

        int firstShapeId = (drawingId * ShapeIdsPerDrawing) + 1;
        int shapeCount = pictures.Count + notes.Count;

        WriteDrawing(writer, sheet, notes, pictures, drawingId, firstShapeId);
        WriteNoteRecords(writer, notes, firstShapeId + pictures.Count);

        // The group shape counts too: it is what the others hang off.
        return shapeCount + 1;
    }

    /// <summary>
    /// The pictures the format can carry, each with its place in the image store. The ones it
    /// cannot are reported rather than dropped silently.
    /// </summary>
    private static List<(Picture, int)> GatherPictures(Worksheet sheet, XlsWriteContext context)
    {
        var pictures = new List<(Picture, int)>(sheet.Pictures.Count);
        foreach (Picture picture in sheet.Pictures)
        {
            if (BlipKind(picture.Format) is not { } kind)
            {
                context.Warn(new WorkbookWarning(
                    "PictureFormatNotWritten",
                    $"'{sheet.Name}' has a {picture.Format} picture, which the .xls drawing layer has no " +
                    "record for; it is not in the file."));
                continue;
            }

            // A BMP is stored as a DIB: the same pixels without the 14-byte file header, which
            // is the one format difference between the file on disk and the record.
            ReadOnlyMemory<byte> content = kind.StripFileHeader ? picture.Content[14..] : picture.Content;
            pictures.Add((picture, context.RegisterBlip(content, picture.Format)));
        }

        return pictures;
    }

    /// <summary>
    /// Writes the drawing itself: the container that holds every shape, then for each shape the
    /// object record that gives it an identity — and, for a note, the text that fills it.
    /// </summary>
    private static void WriteDrawing(
        BiffWriter writer, Worksheet sheet, List<CellComment> notes,
        List<(Picture Picture, int Blip)> pictures, int drawingId, int firstShapeId)
    {
        int spgrLength = GroupShapeSize + (PictureShapeSize * pictures.Count) + (NoteShapeSize * notes.Count);
        int dgLength = 16 + 8 + spgrLength;
        int shapeCount = pictures.Count + notes.Count;

        // The first record opens the containers and carries the group shape and the first shape.
        int start = writer.BeginRecord(BiffRecordType.MsoDrawing);
        WriteHeader(writer, version: 0x0F, instance: 0, DgContainer, dgLength);
        WriteHeader(writer, version: 0x00, instance: drawingId, Dg, 8);
        writer.AppendUInt32((uint)(shapeCount + 1));
        writer.AppendUInt32((uint)(firstShapeId + shapeCount - 1));
        WriteHeader(writer, version: 0x0F, instance: 0, SpgrContainer, spgrLength);
        WriteGroupShape(writer, firstShapeId - 1);

        bool first = true;
        int shapeId = firstShapeId;

        foreach ((Picture picture, int blip) in pictures)
        {
            if (!first)
                start = writer.BeginRecord(BiffRecordType.MsoDrawing);

            WritePictureShape(writer, sheet, picture, blip, shapeId);
            writer.EndRecord(start);
            WritePictureObject(writer, shapeId);
            first = false;
            shapeId++;
        }

        foreach (CellComment note in notes)
        {
            if (!first)
                start = writer.BeginRecord(BiffRecordType.MsoDrawing);

            WriteNoteShape(writer, note, shapeId);
            writer.EndRecord(start);
            WriteObjectAndText(writer, note, shapeId);
            first = false;
            shapeId++;
        }
    }

    /// <summary>Writes the shape every other shape in the drawing belongs to.</summary>
    private static void WriteGroupShape(BiffWriter writer, int shapeId)
    {
        WriteHeader(writer, version: 0x0F, instance: 0, SpContainer, GroupShapeSize - 8);
        WriteHeader(writer, version: 0x01, instance: 0, Spgr, 16);
        writer.AppendUInt32(0);
        writer.AppendUInt32(0);
        writer.AppendUInt32(0x000003FF);            // The group's nominal width and height,
        writer.AppendUInt32(0x000000FF);            // which nothing reads but Excel writes.
        WriteHeader(writer, version: 0x02, instance: 0, Sp, 8);
        writer.AppendUInt32((uint)shapeId);
        writer.AppendUInt32(0x00000005);            // A group, and the drawing's root shape.
    }

    /// <summary>Writes one picture's shape: a frame, the store entry it shows, and its anchor.</summary>
    private static void WritePictureShape(BiffWriter writer, Worksheet sheet, Picture picture, int blip, int shapeId)
    {
        WriteHeader(writer, version: 0x0F, instance: 0, SpContainer, PictureShapeSize - 8);

        WriteHeader(writer, version: 0x02, instance: PictureShapeType, Sp, 8);
        writer.AppendUInt32((uint)shapeId);
        writer.AppendUInt32(0x00000A00);            // Has an anchor, and belongs to a group.

        WriteHeader(writer, version: 0x03, instance: 1, Opt, 6);
        WriteProperty(writer, PibProperty, (uint)blip);

        WritePictureAnchor(writer, sheet, picture);

        WriteHeader(writer, version: 0x00, instance: 0, ClientData, 0);
    }

    /// <summary>Writes the object record that gives a picture shape its identity.</summary>
    private static void WritePictureObject(BiffWriter writer, int shapeId)
    {
        int start = writer.BeginRecord(BiffRecordType.Obj);
        writer.AppendUInt16(0x0015);                // ftCmo: the common properties of an object.
        writer.AppendUInt16(0x0012);
        writer.AppendUInt16(0x0008);                // A picture.
        writer.AppendUInt16(shapeId);
        writer.AppendUInt16(0x6011);                // Locked, printable, and auto-loaded.
        writer.AppendZeros(12);
        writer.AppendUInt16(0);                     // ftEnd.
        writer.AppendUInt16(0);
        writer.EndRecord(start);
    }

    /// <summary>Writes one note's shape: what it is, how it looks, and which cell it hangs on.</summary>
    private static void WriteNoteShape(BiffWriter writer, CellComment note, int shapeId)
    {
        WriteHeader(writer, version: 0x0F, instance: 0, SpContainer, NoteShapeSize - 8);

        WriteHeader(writer, version: 0x02, instance: TextBoxShapeType, Sp, 8);
        writer.AppendUInt32((uint)shapeId);
        writer.AppendUInt32(0x00000A00);            // Has an anchor, and belongs to a group.

        // The look of the yellow box: the pale fill, the thin border, and the flag that keeps
        // it off the printed page. These are the values Excel itself writes.
        WriteHeader(writer, version: 0x03, instance: 10, Opt, 60);
        WriteProperty(writer, 0x0080, 0);           // Text id, which nothing else refers to.
        WriteProperty(writer, 0x0085, 0);           // Text is not wrapped by the shape.
        WriteProperty(writer, 0x0087, 0);
        WriteProperty(writer, 0x0181, 0x08000050);  // Fill: the info-tip colour of the theme.
        WriteProperty(writer, 0x01BF, 0x00010000);  // The fill is real, not just a hit target.
        WriteProperty(writer, 0x01C0, 0x08000040);  // Line: the window-frame colour.
        WriteProperty(writer, 0x01CB, 0x00002535);  // Line width, in EMU.
        WriteProperty(writer, 0x01CE, 0);
        WriteProperty(writer, 0x01FF, 0x00080008);  // The line is drawn, and it is solid.
        WriteProperty(writer, 0x03BF, 0x010A0002);  // Hidden until hovered, and never printed.

        WriteNoteAnchor(writer, note.Cell);

        WriteHeader(writer, version: 0x00, instance: 0, ClientData, 0);
    }

    /// <summary>
    /// Writes where the note's box sits. Excel puts it up and to the right of its cell, two
    /// columns wide and four rows tall, and clamps it to the grid at the edges.
    /// </summary>
    private static void WriteNoteAnchor(BiffWriter writer, CellRef cell)
    {
        int firstColumn = Math.Min(cell.Column, 253);           // Zero-based, plus one to the right.
        int firstRow = Math.Max(cell.Row - 2, 0);
        int lastColumn = Math.Min(firstColumn + 2, 255);
        int lastRow = Math.Min(firstRow + 4, 65535);

        WriteHeader(writer, version: 0x00, instance: 0, ClientAnchor, 18);
        writer.AppendUInt16(0);                                 // The box moves with its cells.
        writer.AppendUInt16(firstColumn);
        writer.AppendUInt16(0);
        writer.AppendUInt16(firstRow);
        writer.AppendUInt16(0);
        writer.AppendUInt16(lastColumn);
        writer.AppendUInt16(0);
        writer.AppendUInt16(lastRow);
        writer.AppendUInt16(0);
    }

    /// <summary>
    /// Writes where a picture sits, as the cells its corners fall in and how far into each —
    /// the exact inverse of the conversion the reader performs, offsets in 1024ths of a
    /// column's width and 256ths of a row's height.
    /// </summary>
    private static void WritePictureAnchor(BiffWriter writer, Worksheet sheet, Picture picture)
    {
        SheetPosition from;
        SheetPosition to;
        PictureAnchor anchor = picture.Anchor;

        switch (anchor.Kind)
        {
            case PictureAnchorKind.TwoCell:
                from = SheetPosition.At(anchor.From, anchor.FromOffset);
                to = SheetPosition.At(anchor.To, anchor.ToOffset);
                break;

            case PictureAnchorKind.OneCell:
                from = SheetPosition.At(anchor.From, anchor.FromOffset);
                to = Advance(sheet, from, anchor.Extent);
                break;

            default:
                from = Locate(sheet, anchor.Position);
                to = Advance(sheet, from, anchor.Extent);
                break;
        }

        int flags = picture.Anchor.MoveBehavior switch
        {
            PictureMoveBehavior.MoveAndResize => 0x0003,
            PictureMoveBehavior.MoveOnly => 0x0001,
            _ => 0x0000,
        };

        WriteHeader(writer, version: 0x00, instance: 0, ClientAnchor, 18);
        writer.AppendUInt16(flags);
        writer.AppendUInt16(Math.Clamp(from.Column - 1, 0, 255));
        writer.AppendUInt16(ColumnFraction(sheet, from));
        writer.AppendUInt16(Math.Clamp(from.Row - 1, 0, 65535));
        writer.AppendUInt16(RowFraction(sheet, from));
        writer.AppendUInt16(Math.Clamp(to.Column - 1, 0, 255));
        writer.AppendUInt16(ColumnFraction(sheet, to));
        writer.AppendUInt16(Math.Clamp(to.Row - 1, 0, 65535));
        writer.AppendUInt16(RowFraction(sheet, to));
    }

    /// <summary>A place on the sheet: a cell and how far into it, in EMU.</summary>
    private readonly record struct SheetPosition(int Row, int Column, long OffsetX, long OffsetY)
    {
        public static SheetPosition At(CellRef cell, (long X, long Y) offset) =>
            new(cell.Row, cell.Column, offset.X, offset.Y);
    }

    /// <summary>
    /// The cell an absolute position falls in, found by walking the grid from its origin —
    /// the only way, because every column on the way may have a width of its own.
    /// </summary>
    private static SheetPosition Locate(Worksheet sheet, (long X, long Y) position)
    {
        int column = 1;
        long x = position.X;
        while (column < 256 && x >= ColumnWidthEmu(sheet, column))
            x -= ColumnWidthEmu(sheet, column++);

        int row = 1;
        long y = position.Y;
        while (row < 65536 && y >= RowHeightEmu(sheet, row))
            y -= RowHeightEmu(sheet, row++);

        return new SheetPosition(row, column, x, y);
    }

    /// <summary>The far corner of a picture, walked out from its near one by its extent.</summary>
    private static SheetPosition Advance(Worksheet sheet, SheetPosition from, (long Width, long Height) extent)
    {
        int column = from.Column;
        long x = from.OffsetX + extent.Width;
        while (column < 256 && x >= ColumnWidthEmu(sheet, column))
            x -= ColumnWidthEmu(sheet, column++);

        int row = from.Row;
        long y = from.OffsetY + extent.Height;
        while (row < 65536 && y >= RowHeightEmu(sheet, row))
            y -= RowHeightEmu(sheet, row++);

        return new SheetPosition(row, column, x, y);
    }

    private static int ColumnFraction(Worksheet sheet, in SheetPosition position)
    {
        long width = ColumnWidthEmu(sheet, position.Column);
        return width <= 0 ? 0 : (int)Math.Clamp(position.OffsetX * 1024 / width, 0, 1023);
    }

    private static int RowFraction(Worksheet sheet, in SheetPosition position)
    {
        long height = RowHeightEmu(sheet, position.Row);
        return height <= 0 ? 0 : (int)Math.Clamp(position.OffsetY * 256 / height, 0, 255);
    }

    /// <summary>
    /// How wide a column is in EMU: characters to pixels the way Excel rounds them, then pixels
    /// to EMU. The same arithmetic the reader runs the other way.
    /// </summary>
    private static long ColumnWidthEmu(Worksheet sheet, int column)
    {
        double characters = sheet.ColumnTable.GetValueOrDefault(column).Width ?? 8.43;
        double pixels = Math.Round((characters * 7.0) + 5.0);
        return (long)(pixels * Emu.PerPixel);
    }

    private static long RowHeightEmu(Worksheet sheet, int row)
    {
        double points = sheet.RowTable.GetValueOrDefault(row).Height ?? 15.0;
        return (long)Math.Round(points * Emu.PerPoint);
    }

    /// <summary>
    /// Writes the object record that gives a note shape an identity, the empty text-box marker
    /// that closes its container, and the text record that fills it.
    /// </summary>
    private static void WriteObjectAndText(BiffWriter writer, CellComment note, int shapeId)
    {
        int objStart = writer.BeginRecord(BiffRecordType.Obj);
        writer.AppendUInt16(0x0015);                // ftCmo: the common properties of an object.
        writer.AppendUInt16(0x0012);
        writer.AppendUInt16(0x0019);                // A note.
        writer.AppendUInt16(shapeId);
        writer.AppendUInt16(0x4011);                // Locked, printable, and auto-filled.
        writer.AppendZeros(12);
        writer.AppendUInt16(0x000D);                // ftNts: the note's own sub-record.
        writer.AppendUInt16(0x0016);
        writer.AppendZeros(22);                     // Its identity guid, which Excel regenerates.
        writer.AppendUInt16(0);                     // ftEnd.
        writer.AppendUInt16(0);
        writer.EndRecord(objStart);

        int textboxStart = writer.BeginRecord(BiffRecordType.MsoDrawing);
        WriteHeader(writer, version: 0x00, instance: 0, ClientTextbox, 0);
        writer.EndRecord(textboxStart);

        WriteText(writer, note.Text);
    }

    /// <summary>
    /// Writes the text of a note. The record itself only states the length; the characters and
    /// the formatting runs follow in continuation records, which is the one shape of this
    /// record every reader expects.
    /// </summary>
    private static void WriteText(BiffWriter writer, string text)
    {
        string clipped = text.Length > 32767 ? text[..32767] : text;

        int start = writer.BeginRecord(BiffRecordType.TxO);
        writer.AppendUInt16(0x0212);                // Left aligned, top aligned, text locked.
        writer.AppendUInt16(0);                     // Not rotated.
        writer.AppendUInt16(0);                     // Reserved, present for this object type.
        writer.AppendUInt32(0);
        writer.AppendUInt16(clipped.Length);
        writer.AppendUInt16(clipped.Length == 0 ? 0 : 16);
        writer.AppendUInt16(0);                     // Font of an empty box.
        writer.AppendUInt16(0);                     // No formula behind the text.
        writer.EndRecord(start);

        if (clipped.Length == 0)
            return;

        int textStart = writer.BeginRecord(BiffRecordType.Continue);
        bool compressible = BiffWriter.IsCompressible(clipped);
        writer.AppendByte(compressible ? (byte)0x00 : (byte)0x01);
        if (compressible)
        {
            foreach (char c in clipped)
                writer.AppendByte((byte)c);
        }
        else
        {
            foreach (char c in clipped)
                writer.AppendUInt16(c);
        }

        writer.EndRecord(textStart);

        // One run covering the whole text, then the marker that ends the run list.
        int runsStart = writer.BeginRecord(BiffRecordType.Continue);
        writer.AppendUInt16(0);                     // First character of the run.
        writer.AppendUInt16(0);                     // Its font, which is the sheet's default.
        writer.AppendUInt32(0);
        writer.AppendUInt16(clipped.Length);
        writer.AppendUInt16(0);
        writer.AppendUInt32(0);
        writer.EndRecord(runsStart);
    }

    /// <summary>Writes the note records that tie the shapes to their cells.</summary>
    private static void WriteNoteRecords(BiffWriter writer, List<CellComment> notes, int firstShapeId)
    {
        for (int i = 0; i < notes.Count; i++)
        {
            CellComment note = notes[i];
            int start = writer.BeginRecord(BiffRecordType.Note);
            writer.AppendUInt16(note.Cell.Row - 1);
            writer.AppendUInt16(note.Cell.Column - 1);
            writer.AppendUInt16(0);                 // Shown only when the cell is hovered.
            writer.AppendUInt16(firstShapeId + i);
            writer.AppendUnicodeString(Author(note));
            writer.AppendByte(0);                   // Unused, but the record is short without it.
            writer.EndRecord(start);
        }
    }

    /// <summary>
    /// The author of a note, cut to what the format allows. An empty one is not allowed at all,
    /// so a note that never had an author gets a space.
    /// </summary>
    private static string Author(CellComment note) => note.Author switch
    {
        null or "" => " ",
        { Length: > 54 } long_ => long_[..54],
        { } author => author,
    };

    /// <summary>
    /// Writes the workbook-wide drawing record: how many shapes exist, how their ids were
    /// handed out, the image store, and the palette the drawing layer offers in its own
    /// colour picker.
    /// </summary>
    /// <param name="writer">Global substream being written.</param>
    /// <param name="context">The write context, which holds the drawings and the image store.</param>
    public static void WriteDrawingGroup(BiffWriter writer, XlsWriteContext context)
    {
        IReadOnlyList<(int DrawingId, int ShapeCount)> drawings = context.Drawings;
        if (drawings.Count == 0)
            return;

        int totalShapes = 0;
        int highestShapeId = 0;
        foreach ((int drawingId, int shapeCount) in drawings)
        {
            totalShapes += shapeCount;
            highestShapeId = Math.Max(highestShapeId, (drawingId * ShapeIdsPerDrawing) + shapeCount - 1);
        }

        int identifierBlock = 16 + (8 * drawings.Count);
        int storeLength = StoreLength(context.Blips);
        int containerLength = 8 + identifierBlock + (storeLength > 0 ? 8 + storeLength : 0) + 8 + 18 + 8 + 16;

        int start = writer.BeginRecord(BiffRecordType.MsoDrawingGroup);
        WriteHeader(writer, version: 0x0F, instance: 0, DggContainer, containerLength);
        WriteHeader(writer, version: 0x00, instance: 0, Dgg, identifierBlock);
        writer.AppendUInt32((uint)(highestShapeId + 1));
        writer.AppendUInt32((uint)(drawings.Count + 1));
        writer.AppendUInt32((uint)totalShapes);
        writer.AppendUInt32((uint)drawings.Count);
        foreach ((int drawingId, int shapeCount) in drawings)
        {
            writer.AppendUInt32((uint)drawingId);
            writer.AppendUInt32((uint)shapeCount);
        }

        WriteBlipStore(writer, context.Blips, storeLength);

        WriteHeader(writer, version: 0x03, instance: 3, Opt, 18);
        WriteProperty(writer, 0x00BF, 0x00080008);      // Text: no auto text margin.
        WriteProperty(writer, 0x0181, 0x08000041);      // Default fill colour.
        WriteProperty(writer, 0x01C0, 0x08000040);      // Default line colour.

        WriteHeader(writer, version: 0x00, instance: 4, SplitMenuColors, 16);
        writer.AppendUInt32(0x0800000D);
        writer.AppendUInt32(0x0800000C);
        writer.AppendUInt32(0x08000017);
        writer.AppendUInt32(0x100000F7);
        writer.EndRecord(start);
    }

    /// <summary>How a picture format is spelled in the store, or null when it cannot be.</summary>
    /// <remarks>
    /// GIF and TIFF have no place here — the format predates the one and never adopted the
    /// other in practice — and an unknown format has no record type to claim.
    /// </remarks>
    private static (ushort RecordType, int Instance, byte BlipType, bool Tagged, bool StripFileHeader)? BlipKind(
        PictureFormat format) => format switch
    {
        PictureFormat.Png => ((ushort)0xF01E, 0x6E0, (byte)6, true, false),
        PictureFormat.Jpeg => ((ushort)0xF01D, 0x46A, (byte)5, true, false),
        PictureFormat.Bmp => ((ushort)0xF01F, 0x7A8, (byte)7, true, true),
        PictureFormat.Emf => ((ushort)0xF01A, 0x3D4, (byte)2, false, false),
        PictureFormat.Wmf => ((ushort)0xF01B, 0x216, (byte)3, false, false),
        _ => null,
    };

    /// <summary>The size of one store entry's body, header of the embedded record included.</summary>
    private static int EntryLength(ReadOnlyMemory<byte> content, PictureFormat format)
    {
        (_, _, _, bool tagged, _) = BlipKind(format)!.Value;

        // The BSE body is 36 bytes, then the embedded record: its 8-byte header, one digest,
        // and either a tag byte or the 34-byte metafile header before the image itself.
        return 36 + 8 + 16 + (tagged ? 1 : 34) + content.Length;
    }

    private static int StoreLength(
        IReadOnlyList<(ReadOnlyMemory<byte> Content, PictureFormat Format, int References)> blips)
    {
        if (blips.Count == 0)
            return 0;

        int length = 0;
        foreach ((ReadOnlyMemory<byte> content, PictureFormat format, _) in blips)
            length += 8 + EntryLength(content, format);

        return length;
    }

    /// <summary>Writes the image store: one entry per distinct image, in the order they were claimed.</summary>
    private static void WriteBlipStore(
        BiffWriter writer,
        IReadOnlyList<(ReadOnlyMemory<byte> Content, PictureFormat Format, int References)> blips,
        int storeLength)
    {
        if (blips.Count == 0)
            return;

        WriteHeader(writer, version: 0x0F, instance: blips.Count, BStoreContainer, storeLength);

        foreach ((ReadOnlyMemory<byte> content, PictureFormat format, int references) in blips)
        {
            (ushort recordType, int instance, byte blipType, bool tagged, _) = BlipKind(format)!.Value;

            // The digest identifies the image for dedup; readers take it on trust rather than
            // recomputing it, so any digest of the bytes serves as long as equal bytes get
            // equal digests.
            byte[] digest = System.Security.Cryptography.MD5.HashData(content.Span);

            int blipLength = 16 + (tagged ? 1 : 34) + content.Length;

            WriteHeader(writer, version: 0x02, instance: blipType, Bse, 36 + 8 + blipLength);
            writer.AppendByte(blipType);            // btWin32.
            writer.AppendByte(blipType);            // btMacOS: the same picture either way.
            writer.Append(digest);
            writer.AppendUInt16(0xFF);              // Tag.
            writer.AppendUInt32((uint)(8 + blipLength));
            writer.AppendUInt32((uint)references);
            writer.AppendUInt32(0);                 // foDelay: the image is right here.
            writer.AppendByte(0);                   // Usage.
            writer.AppendByte(0);                   // Name length: these entries carry none.
            writer.AppendUInt16(0);                 // Unused pair.

            WriteHeader(writer, version: 0x00, instance: instance, recordType, blipLength);
            writer.Append(digest);

            if (tagged)
            {
                writer.AppendByte(0xFF);
            }
            else
            {
                // The metafile header: sizes, nominal bounds, and the flag that says the bytes
                // follow uncompressed.
                writer.AppendUInt32((uint)content.Length);
                writer.AppendZeros(16);             // rcBounds.
                writer.AppendUInt32(0);             // ptSize.x.
                writer.AppendUInt32(0);             // ptSize.y.
                writer.AppendUInt32((uint)content.Length);
                writer.AppendByte(0xFE);            // Stored as-is.
                writer.AppendByte(0xFE);            // No filter.
            }

            writer.Append(content.Span);
        }
    }

    /// <summary>Writes an Office Art record header: a version, an instance, a type and a length.</summary>
    private static void WriteHeader(BiffWriter writer, int version, int instance, ushort type, int length)
    {
        writer.AppendUInt16((instance << 4) | (version & 0x0F));
        writer.AppendUInt16(type);
        writer.AppendUInt32((uint)length);
    }

    /// <summary>Writes one shape property: its id and its value.</summary>
    private static void WriteProperty(BiffWriter writer, int id, uint value)
    {
        writer.AppendUInt16(id);
        writer.AppendUInt32(value);
    }
}
