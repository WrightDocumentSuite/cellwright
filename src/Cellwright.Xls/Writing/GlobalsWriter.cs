using Cellwright.Styles;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the global substream: the records that describe the workbook as a whole, in the order
/// the format requires them.
/// </summary>
/// <remarks>
/// Two things here are written blind and fixed up afterwards. The BoundSheet8 records carry the
/// byte offset of each sheet's substream, which is not known until the sheets have been
/// written; and the shared string table is built by walking every sheet first, so the cells
/// have to be visited before the table that they index into can be emitted.
/// </remarks>
internal sealed class GlobalsWriter(XlsWriteContext context, StyleTable styles)
{
    /// <summary>Offsets of the BoundSheet8 position fields, to be patched once sheets are placed.</summary>
    private readonly List<int> _sheetOffsetFields = [];

    /// <summary>Where each sheet's BoundSheet8 offset field sits in the stream.</summary>
    public IReadOnlyList<int> SheetOffsetFields => _sheetOffsetFields;

    /// <summary>Writes the whole global substream.</summary>
    public void Write(BiffWriter writer)
    {
        WriteBof(writer);
        WriteHeaderRecords(writer);
        styles.Write(writer);
        WritePalette(writer);
        WriteBoundSheets(writer);

        // The names are encoded before the supporting-link table is written because encoding
        // one may add an entry to it, and the format puts that table first.
        List<(DefinedName Name, byte[] Rgce)> names = EncodeDefinedNames();
        WriteSupportingLinks(writer);
        WriteDefinedNames(writer, names);
        EscherDrawingWriter.WriteDrawingGroup(writer, context);
        WriteSharedStrings(writer);
        writer.WriteRecord(BiffRecordType.Eof);
    }

    private static void WriteBof(BiffWriter writer)
    {
        Span<byte> body = stackalloc byte[16];
        body.Clear();
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, 0x0600);   // BIFF8.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[2..], 0x0005); // Globals.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[4..], 0x0DBB); // Build.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[6..], 0x07CC); // Build year.
        writer.WriteRecord(BiffRecordType.Bof, body);
    }

    private void WriteHeaderRecords(BiffWriter writer)
    {
        // InterfaceHdr through CodePage: Excel refuses a file that skips them, even though
        // nothing in them varies for a workbook written from scratch.
        WriteUInt16Record(writer, BiffRecordType.Interfacehdr, 0x04B0);   // ANSI code page.
        WriteMms(writer);
        writer.WriteRecord(0x00E2, []);                                   // InterfaceEnd.
        WriteWriteAccess(writer);
        WriteUInt16Record(writer, BiffRecordType.CodePage, 0x04B0);       // UTF-16.

        WriteUInt16Record(writer, 0x0161, 0x0000);                        // DSF: not a double stream file.
        writer.WriteRecord(0x01C0);                                       // Excel9File, which has no body.
        WriteTabIds(writer);

        WriteBoolRecord(writer, 0x000C, 0x0064);                          // CalcCount.
        WriteBoolRecord(writer, 0x000D, 0x0001);                          // CalcMode: automatic.
        WriteBoolRecord(writer, 0x000E, 0x0001);                          // Precision as displayed off.
        WriteBoolRecord(writer, 0x000F, 0x0000);                          // RefMode: A1.
        WriteDelta(writer);
        WriteBoolRecord(writer, 0x0011, 0x0000);                          // Iteration off.
        WriteBoolRecord(writer, BiffRecordType.Date1904,
            context.Workbook.DateSystem == ExcelDateSystem.Date1904 ? 1 : 0);
        WriteBoolRecord(writer, 0x005E, 0x0000);                          // BackupFile off.
        WriteBoolRecord(writer, 0x008D, 0x0000);                          // HideObj: everything shown.

        WriteWindowProtection(writer);
    }

    /// <summary>Writes the two menu counts that Excel has kept at zero since version 5.</summary>
    private static void WriteMms(BiffWriter writer)
    {
        Span<byte> body = stackalloc byte[2];
        body.Clear();
        writer.WriteRecord(0x00C1, body);
    }

    /// <summary>Writes the sheet identifiers, which revision tracking numbers the tabs by.</summary>
    private void WriteTabIds(BiffWriter writer)
    {
        int start = writer.BeginRecord(0x013D);
        for (int i = 0; i < context.Workbook.Sheets.Count; i++)
            writer.AppendUInt16(i + 1);
        writer.EndRecord(start);
    }

    /// <summary>Writes the convergence threshold of iterative calculation, Excel's own default.</summary>
    private static void WriteDelta(BiffWriter writer)
    {
        int start = writer.BeginRecord(0x0010);
        writer.AppendDouble(0.001);
        writer.EndRecord(start);
    }

    /// <summary>
    /// Writes the name of whoever last saved the file. The record is a fixed 112 bytes however
    /// short the name is, and two spaces are how the format says nobody is named.
    /// </summary>
    private static void WriteWriteAccess(BiffWriter writer)
    {
        const int RecordSize = 112;
        int start = writer.BeginRecord(0x005C);
        writer.AppendUnicodeString("  ");
        writer.AppendZeros(RecordSize - (writer.Length - start));
        writer.EndRecord(start);
    }

    private void WriteWindowProtection(BiffWriter writer)
    {
        WorkbookProtection protection = context.Workbook.Protection;
        WriteBoolRecord(writer, 0x0012, protection.LockStructure ? 1 : 0);   // Protect.
        WriteBoolRecord(writer, 0x0019, protection.LockWindows ? 1 : 0);     // WinProtect.
        WriteBoolRecord(writer, 0x00DD, 0);                                  // ScenProtect.

        // Window1: the workbook's window geometry, which Excel needs but does not mind being
        // told the defaults for.
        Span<byte> window = stackalloc byte[18];
        window.Clear();
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(window[8..], 0x0038);  // Flags.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(window[12..], 1);      // Tabs shown.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(window[14..], 1);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(window[16..], 0x0258); // Tab bar width.
        writer.WriteRecord(BiffRecordType.Window1, window);
    }

    /// <summary>
    /// Writes the colour palette. Every colour the workbook uses is quantised into the 56 slots
    /// the format has, since a cell can only name a slot and not a value.
    /// </summary>
    private void WritePalette(BiffWriter writer)
    {
        IReadOnlyList<uint> palette = styles.Palette;
        if (palette.Count == 0)
            return;

        int start = writer.BeginRecord(BiffRecordType.Palette);
        writer.AppendUInt16(palette.Count);
        foreach (uint color in palette)
        {
            writer.AppendByte((byte)((color >> 16) & 0xFF));
            writer.AppendByte((byte)((color >> 8) & 0xFF));
            writer.AppendByte((byte)(color & 0xFF));
            writer.AppendByte(0);
        }

        writer.EndRecord(start);
    }

    private void WriteBoundSheets(BiffWriter writer)
    {
        foreach (Worksheet sheet in context.Workbook.Sheets)
        {
            int start = writer.BeginRecord(BiffRecordType.BoundSheet8);
            _sheetOffsetFields.Add(writer.ReserveUInt32());
            writer.AppendByte(sheet.Visibility switch
            {
                SheetVisibility.Hidden => 0x01,
                SheetVisibility.VeryHidden => 0x02,
                _ => 0x00,
            });
            writer.AppendByte(0x00); // A worksheet; this library models nothing else.
            writer.AppendShortUnicodeString(sheet.Name);
            writer.EndRecord(start);
        }
    }

    /// <summary>
    /// Writes the supporting-link table that every sheet-qualified reference points into. A
    /// workbook that only refers to itself still needs one entry saying so, plus a row in the
    /// reference table for each sheet named by a formula.
    /// </summary>
    private void WriteSupportingLinks(BiffWriter writer)
    {
        IReadOnlyList<int> targets = context.ExternSheetTargets;
        if (targets.Count == 0)
            return;

        int supBook = writer.BeginRecord(BiffRecordType.SupBook);
        writer.AppendUInt16(context.Workbook.Sheets.Count);
        writer.AppendUInt16(0x0401);            // The link points at this workbook.
        writer.EndRecord(supBook);

        int externSheet = writer.BeginRecord(BiffRecordType.ExternSheet);
        writer.AppendUInt16(targets.Count);
        foreach (int sheet in targets)
        {
            writer.AppendUInt16(0);             // The only supporting link, written just above.
            writer.AppendUInt16(sheet);
            writer.AppendUInt16(sheet);
        }

        writer.EndRecord(externSheet);
    }

    /// <summary>Encodes every defined name, dropping the ones the format cannot carry.</summary>
    private List<(DefinedName Name, byte[] Rgce)> EncodeDefinedNames()
    {
        var encoder = new PtgEncoder(context);
        var encoded = new List<(DefinedName, byte[])>(context.Workbook.Names.Count);
        int index = 0;

        // The names have to be registered before any is encoded: one name may refer to another,
        // and the reference is by position in this very table.
        foreach (DefinedName name in context.Workbook.Names)
            context.RegisterName(name.Name, index++);

        foreach (DefinedName name in context.Workbook.Names)
        {
            byte[]? rgce = encoder.Encode(name.RefersTo, new CellRef(1, 1));
            if (rgce is null)
            {
                context.Warn(new WorkbookWarning(
                    "UnwritableName",
                    $"Defined name '{name.Name}' refers to something the .xls format has no encoding for and was left out."));
                continue;
            }

            encoded.Add((name, rgce));
        }

        return encoded;
    }

    private void WriteDefinedNames(BiffWriter writer, List<(DefinedName Name, byte[] Rgce)> names)
    {
        foreach ((DefinedName name, byte[] rgce) in names)
        {
            int start = writer.BeginRecord(BiffRecordType.Lbl);
            writer.AppendUInt16(name.IsHidden ? 0x0001 : 0x0000);
            writer.AppendByte(0);                                   // No keyboard shortcut.
            writer.AppendByte((byte)Math.Min(name.Name.Length, 255));
            writer.AppendUInt16(rgce.Length);
            writer.AppendUInt16(0);                                 // Reserved.
            writer.AppendUInt16(SheetScopeOf(name));
            writer.AppendZeros(4);                                  // No help or status text.
            writer.AppendByte(BiffWriter.IsCompressible(name.Name) ? (byte)0x00 : (byte)0x01);
            AppendNameCharacters(writer, name.Name);
            writer.Append(rgce);
            writer.EndRecord(start);
        }
    }

    private static void AppendNameCharacters(BiffWriter writer, string name)
    {
        if (BiffWriter.IsCompressible(name))
        {
            foreach (char c in name)
                writer.AppendByte((byte)c);
            return;
        }

        foreach (char c in name)
            writer.AppendUInt16(c);
    }

    /// <summary>Returns the one-based sheet a name is local to, or 0 for a workbook-wide one.</summary>
    private int SheetScopeOf(DefinedName name)
    {
        if (name.Scope is null)
            return 0;

        for (int i = 0; i < context.Workbook.Sheets.Count; i++)
        {
            if (ReferenceEquals(context.Workbook.Sheets[i], name.Scope))
                return i + 1;
        }

        return 0;
    }

    /// <summary>
    /// Writes the shared string table. However long it gets it is one logical record, split
    /// across continuations — and a string may straddle the split, in which case the piece that
    /// carries the rest of it opens by restating whether the characters are one byte or two.
    /// </summary>
    private void WriteSharedStrings(BiffWriter writer)
    {
        IReadOnlyList<string> strings = context.SharedStrings;
        if (strings.Count == 0)
            return;

        IReadOnlyList<(ushort CharPos, ushort FontIndex)[]?> runs = context.SharedStringRuns;
        var pieces = new SstPieceWriter(writer);
        pieces.AppendUInt32((uint)context.TotalStringReferences);
        pieces.AppendUInt32((uint)strings.Count);
        for (int i = 0; i < strings.Count; i++)
            pieces.AppendString(strings[i], i < runs.Count ? runs[i] : null);
        pieces.Flush();

        WriteExtSst(writer, strings.Count);
    }

    /// <summary>
    /// Lays the shared string table out record by record. It fills each record to the byte and
    /// starts a continuation when it runs out, which is what Excel does and therefore what
    /// every reader of these files is built to expect.
    /// </summary>
    private struct SstPieceWriter(BiffWriter writer)
    {
        private readonly List<byte> _current = new(BiffReader.MaxRecordSize);
        private bool _started;

        /// <summary>Appends a four-byte header field, which never straddles a record.</summary>
        public void AppendUInt32(uint value)
        {
            Span<byte> bytes = stackalloc byte[4];
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(bytes, value);
            foreach (byte b in bytes)
                _current.Add(b);
        }

        /// <summary>Appends one string, splitting its characters across records if it has to.</summary>
        public void AppendString(string text, (ushort CharPos, ushort FontIndex)[]? runs = null)
        {
            bool compressed = BiffWriter.IsCompressible(text);
            int width = compressed ? 1 : 2;
            bool hasRuns = runs is { Length: > 0 };
            int headerBytes = hasRuns ? 5 : 3;

            // The prefix is never split: a record that cannot hold it plus at least
            // one character ends here instead.
            if (Remaining < headerBytes + width)
                StartContinuation(flagsByte: null);

            _current.Add((byte)(text.Length & 0xFF));
            _current.Add((byte)((text.Length >> 8) & 0xFF));
            byte flags = (byte)((compressed ? 0x00 : 0x01) | (hasRuns ? 0x08 : 0x00));
            _current.Add(flags);
            if (hasRuns)
            {
                _current.Add((byte)(runs!.Length & 0xFF));
                _current.Add((byte)((runs.Length >> 8) & 0xFF));
            }

            int written = 0;
            while (written < text.Length)
            {
                int take = Math.Min(text.Length - written, Remaining / width);
                if (take == 0)
                {
                    StartContinuation(compressed ? (byte)0x00 : (byte)0x01);
                    continue;
                }

                for (int i = 0; i < take; i++)
                {
                    char c = text[written + i];
                    _current.Add((byte)(c & 0xFF));
                    if (!compressed)
                        _current.Add((byte)((c >> 8) & 0xFF));
                }

                written += take;
            }

            if (hasRuns)
            {
                foreach ((ushort charPos, ushort fontIndex) in runs!)
                {
                    if (Remaining < 4)
                        StartContinuation(flagsByte: null);

                    _current.Add((byte)(charPos & 0xFF));
                    _current.Add((byte)((charPos >> 8) & 0xFF));
                    _current.Add((byte)(fontIndex & 0xFF));
                    _current.Add((byte)((fontIndex >> 8) & 0xFF));
                }
            }
        }

        /// <summary>Writes the last piece out.</summary>
        public void Flush()
        {
            if (_current.Count == 0)
                return;

            writer.WriteRecord(_started ? BiffRecordType.Continue : BiffRecordType.Sst, [.. _current]);
            _current.Clear();
            _started = true;
        }

        private readonly int Remaining => BiffReader.MaxRecordSize - _current.Count;

        /// <summary>
        /// Closes the current record and opens a continuation, restating the encoding when the
        /// break falls inside a string's characters.
        /// </summary>
        private void StartContinuation(byte? flagsByte)
        {
            Flush();
            if (flagsByte is { } flags)
                _current.Add(flags);
        }
    }

    /// <summary>
    /// Writes the ExtSST index. It is a lookup table into the SST that Excel builds to jump to a
    /// string without reading the ones before it; an empty one is accepted, and the file is
    /// correct either way, so the buckets are left unpopulated.
    /// </summary>
    private static void WriteExtSst(BiffWriter writer, int stringCount)
    {
        int perBucket = Math.Max(8, (stringCount / 128) + 1);
        Span<byte> body = stackalloc byte[2];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, (ushort)perBucket);
        writer.WriteRecord(BiffRecordType.ExtSst, body);
    }

    private static void WriteUInt16Record(BiffWriter writer, ushort type, int value)
    {
        Span<byte> body = stackalloc byte[2];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, (ushort)value);
        writer.WriteRecord(type, body);
    }

    private static void WriteBoolRecord(BiffWriter writer, ushort type, int value) =>
        WriteUInt16Record(writer, type, value);
}
