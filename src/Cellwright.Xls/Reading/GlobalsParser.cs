using System.Buffers.Binary;
using System.Text;
using Cellwright.Styles;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads the global substream that opens a workbook: which sheets it has and where they start,
/// the shared strings every text cell points into, the style tables, and the defined names.
/// </summary>
internal sealed class GlobalsParser(BiffReader reader, StylePool stylePool, Action<WorkbookWarning> warn)
{
    private readonly List<XlsFontRecord> _fonts = [];
    private readonly Dictionary<int, string> _formats = [];
    private readonly List<uint> _palette = [];

    /// <summary>
    /// The workbook's image store, in pieces. A drawing group larger than one record does not
    /// continue in a <c>Continue</c> record the reader would join for us: it continues in
    /// another <c>MsoDrawingGroup</c>, so the pieces are gathered here and joined at the end.
    /// Parsing them one at a time reads only as much of the store as the first piece held.
    /// </summary>
    private readonly List<byte[]> _drawingGroup = [];

    /// <summary>Reads from the current position to the end of the global substream.</summary>
    /// <param name="version">Version the opening BOF declared.</param>
    public XlsGlobals Parse(BiffVersion version)
    {
        var globals = new XlsGlobals { Version = version };
        var rawFormats = new List<XlsXfRecord>();

        while (reader.Read())
        {
            ReadOnlySpan<byte> body = reader.Record.Span;
            switch (reader.RecordType)
            {
                case BiffRecordType.Eof:
                    Finish(globals, rawFormats);
                    return globals;

                case BiffRecordType.Bof:
                    // The global substream should close with EOF; when it does not, the first
                    // sheet's BOF is where it ends, and reading on would take that sheet's
                    // records for workbook-wide ones.
                    warn(new WorkbookWarning(
                        "MissingEof",
                        "The workbook's global records have no end marker; they were taken to end where the first sheet begins."));
                    Finish(globals, rawFormats);
                    return globals;

                case BiffRecordType.CodePage:
                    globals.Encoding = ResolveEncoding(body);
                    break;

                case BiffRecordType.Date1904:
                    if (body.Length >= 2 && BinaryPrimitives.ReadUInt16LittleEndian(body) != 0)
                        globals.DateSystem = ExcelDateSystem.Date1904;
                    break;

                case BiffRecordType.BoundSheet8:
                    ReadBoundSheet(globals, body);
                    break;

                case BiffRecordType.Sst:
                    globals.SharedStrings = ReadSharedStrings(body);
                    break;

                case BiffRecordType.Font or BiffRecordType.OldFont:
                    _fonts.Add(ReadFont(body, globals));
                    break;

                case BiffRecordType.Format or BiffRecordType.OldFormat:
                    ReadFormat(body, globals);
                    break;

                case BiffRecordType.Xf or BiffRecordType.Biff2Xf or BiffRecordType.Biff3Xf or BiffRecordType.Biff4Xf:
                    rawFormats.Add(ReadXf(body, globals.Version));
                    break;

                case BiffRecordType.Palette:
                    ReadPalette(body);
                    break;

                // Every image in the file is stored once, here, and the sheets point at them.
                case BiffRecordType.MsoDrawingGroup:
                    _drawingGroup.Add(body.ToArray());
                    break;

                case BiffRecordType.Lbl:
                    ReadDefinedName(globals, body);
                    break;

                case BiffRecordType.SupBook:
                    globals.SupBooks.Add(ReadSupBook(body));
                    break;

                case BiffRecordType.ExternName:
                    ReadExternName(globals, body);
                    break;

                case BiffRecordType.ExternSheet:
                    ReadExternSheet(globals, body);
                    break;

                default:
                    break;
            }
        }

        Finish(globals, rawFormats);
        return globals;
    }

    /// <summary>
    /// Turns the raw XF table into styles once everything it refers to — fonts, formats and the
    /// palette — has been read, since a file may state them in any order.
    /// </summary>
    private void Finish(XlsGlobals globals, List<XlsXfRecord> rawFormats)
    {
        globals.Palette.AddRange(_palette);
        if (_drawingGroup.Count > 0)
            globals.Blips.AddRange(EscherPictures.ReadStore(EscherPictures.Join(_drawingGroup)));

        foreach (XlsXfRecord xf in rawFormats)
        {
            string numberFormat = ResolveFormatCode(xf.FormatIndex);
            CellStyle style = BuildStyle(xf, numberFormat);
            globals.CellFormats.Add(new StyleId(stylePool.Intern(style)));
            globals.TemporalKinds.Add(NumberFormatClassifier.Classify(xf.FormatIndex, numberFormat.Length == 0 ? null : numberFormat));
        }
    }

    private string ResolveFormatCode(int formatIndex)
    {
        if (_formats.TryGetValue(formatIndex, out string? custom))
            return custom;

        // Builtin ids mean the same thing in both formats, so the xlsx table answers for both.
        return BuiltinNumberFormats.GetCode(formatIndex) ?? "";
    }

    private CellStyle BuildStyle(in XlsXfRecord xf, string numberFormat)
    {
        XlsFontRecord font = (uint)xf.FontIndex < (uint)_fonts.Count ? _fonts[xf.FontIndex] : XlsFontRecord.Default;
        return new CellStyle
        {
            NumberFormat = numberFormat,
            Font = new FontStyle
            {
                Name = font.Name,
                Size = font.Size,
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = font.Underline,
                Strikethrough = font.Strikethrough,
                VerticalAlign = font.VerticalAlign,
                Color = ResolveColor(font.ColorIndex),
                Charset = font.Charset,
                Family = font.Family,
            },
            Fill = xf.FillPattern == FillPattern.None
                ? FillStyle.None
                : FillStyle.Patterned(xf.FillPattern, ResolveColor(xf.ForegroundColorIndex), ResolveColor(xf.BackgroundColorIndex)),
            Border = new BorderStyle
            {
                Left = new BorderEdge(xf.Left, ResolveColor(xf.LeftColorIndex)),
                Right = new BorderEdge(xf.Right, ResolveColor(xf.RightColorIndex)),
                Top = new BorderEdge(xf.Top, ResolveColor(xf.TopColorIndex)),
                Bottom = new BorderEdge(xf.Bottom, ResolveColor(xf.BottomColorIndex)),
                Diagonal = new BorderEdge(xf.Diagonal, ResolveColor(xf.DiagonalColorIndex)),
                DiagonalUp = xf.DiagonalUp,
                DiagonalDown = xf.DiagonalDown,
            },
            Alignment = new Alignment
            {
                Horizontal = xf.Horizontal,
                Vertical = xf.Vertical,
                WrapText = xf.WrapText,
                Indent = xf.Indent,
                TextRotation = xf.TextRotation,
                ShrinkToFit = xf.ShrinkToFit,
                ReadingOrder = xf.ReadingOrder,
            },
            Protection = new CellProtection { Locked = xf.Locked, Hidden = xf.FormulaHidden },
        };
    }

    /// <summary>
    /// Resolves a palette index to a color. The legacy format has no notion of a theme: every
    /// color is a slot in a table of 56 the workbook may override, plus a few fixed meanings.
    /// </summary>
    private ExcelColor ResolveColor(int index)
    {
        // 0x7FFF is "automatic"; 64 and 65 are the system window colors, which have no value.
        if (index is 0x7FFF or 64 or 65 or 0)
            return ExcelColor.Auto;

        int slot = index - 8;
        if ((uint)slot < (uint)_palette.Count)
            return ExcelColor.FromArgb(_palette[slot]);

        return ExcelColor.FromIndexed(index);
    }

    private void ReadPalette(ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        _palette.Clear();
        int count = BinaryPrimitives.ReadUInt16LittleEndian(body);
        for (int i = 0; i < count && 2 + (i * 4) + 4 <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(2 + (i * 4), 4);
            _palette.Add(0xFF000000u | ((uint)entry[0] << 16) | ((uint)entry[1] << 8) | entry[2]);
        }
    }

    private void ReadBoundSheet(XlsGlobals globals, ReadOnlySpan<byte> body)
    {
        if (body.Length < 6)
            return;

        int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        byte state = body[4];
        byte type = body[5];
        int cursor = 6;
        string name = globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadShortUnicodeString(body, ref cursor)
            : BiffStrings.ReadByteString(body, ref cursor, globals.Encoding, lengthBytes: 1);

        SheetVisibility visibility = (state & 0x03) switch
        {
            1 => SheetVisibility.Hidden,
            2 => SheetVisibility.VeryHidden,
            _ => SheetVisibility.Visible,
        };

        XlsSheetKind kind = type switch
        {
            0x00 => XlsSheetKind.Worksheet,
            0x02 => XlsSheetKind.Chart,
            _ => XlsSheetKind.Other,
        };

        globals.Sheets.Add(new XlsSheetEntry(name, offset, visibility, kind));
    }

    private string[] ReadSharedStrings(ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return [];

        int unique = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]);
        if (unique <= 0)
            return [];

        // A corrupt count could ask for gigabytes; the record's own length bounds what can
        // possibly be in it, since the shortest string costs three bytes.
        unique = Math.Min(unique, (body.Length / 3) + 1);

        var strings = new string[unique];
        var cursor = new BiffStrings.SstCursor(body, reader.ContinuationSizes, offset: 8);
        int read = 0;
        while (read < unique && cursor.HasMore)
            strings[read++] = cursor.ReadNext();

        if (read == unique)
            return strings;

        warn(new WorkbookWarning(
            "TruncatedSst",
            $"The shared string table promises {unique} strings but holds {read}; the rest read as empty."));
        return strings;
    }

    private XlsFontRecord ReadFont(ReadOnlySpan<byte> body, XlsGlobals globals)
    {
        if (body.Length < 4)
            return XlsFontRecord.Default;

        double size = BinaryPrimitives.ReadUInt16LittleEndian(body) / 20.0;
        ushort attributes = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        bool italic = (attributes & 0x0002) != 0;
        bool strikethrough = (attributes & 0x0008) != 0;

        if (globals.Version <= BiffVersion.Biff4 || body.Length < 14)
        {
            int shortCursor = 4;
            string shortName = BiffStrings.ReadByteString(body, ref shortCursor, globals.Encoding, lengthBytes: 1);
            return XlsFontRecord.Default with
            {
                Name = shortName.Length == 0 ? XlsFontRecord.Default.Name : shortName,
                Size = size,
                Bold = (attributes & 0x0001) != 0,
                Italic = italic,
                Strikethrough = strikethrough,
            };
        }

        int colorIndex = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int weight = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        int escapement = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);
        byte underline = body[10];
        byte family = body[11];
        byte charset = body[12];

        int cursor = 14;
        string name = globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadShortUnicodeString(body, ref cursor)
            : BiffStrings.ReadByteString(body, ref cursor, globals.Encoding, lengthBytes: 1);

        return new XlsFontRecord
        {
            Name = name.Length == 0 ? XlsFontRecord.Default.Name : name,
            Size = size,
            Bold = weight >= 700,
            Italic = italic,
            Strikethrough = strikethrough,
            Underline = underline switch
            {
                0x01 => UnderlineStyle.Single,
                0x02 => UnderlineStyle.Double,
                0x21 => UnderlineStyle.SingleAccounting,
                0x22 => UnderlineStyle.DoubleAccounting,
                _ => UnderlineStyle.None,
            },
            VerticalAlign = escapement switch
            {
                1 => FontVerticalAlign.Superscript,
                2 => FontVerticalAlign.Subscript,
                _ => FontVerticalAlign.Baseline,
            },
            ColorIndex = colorIndex,
            Family = family,
            Charset = charset,
        };
    }

    private void ReadFormat(ReadOnlySpan<byte> body, XlsGlobals globals)
    {
        if (globals.Version <= BiffVersion.Biff4)
        {
            // Before BIFF5 the record carries no id: formats are numbered by arrival.
            int legacyCursor = 0;
            string legacyCode = BiffStrings.ReadByteString(body, ref legacyCursor, globals.Encoding, lengthBytes: 1);
            _formats[_formats.Count] = legacyCode;
            return;
        }

        if (body.Length < 2)
            return;

        int id = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int cursor = 2;
        string code = globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadUnicodeString(body, ref cursor)
            : BiffStrings.ReadByteString(body, ref cursor, globals.Encoding, lengthBytes: 1);
        _formats[id] = code;
    }

    private static XlsXfRecord ReadXf(ReadOnlySpan<byte> body, BiffVersion version) =>
        version >= BiffVersion.Biff8 && body.Length >= 20
            ? XlsXfRecord.ParseBiff8(body)
            : XlsXfRecord.ParseLegacy(body, version);

    private void ReadDefinedName(XlsGlobals globals, ReadOnlySpan<byte> body)
    {
        if (body.Length < 14)
            return;

        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int nameLength = body[3];
        int rgceLength = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int sheetIndex = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);
        bool isBuiltIn = (flags & 0x0020) != 0;
        bool isHidden = (flags & 0x0001) != 0;

        int cursor = 14;
        string name;
        if (globals.Version >= BiffVersion.Biff8)
        {
            // The name here has a length that was already read, so only the flags byte and the
            // characters remain.
            bool wide = cursor < body.Length && (body[cursor] & 0x01) != 0;
            cursor++;
            int byteCount = Math.Min(wide ? nameLength * 2 : nameLength, Math.Max(0, body.Length - cursor));
            name = wide
                ? Encoding.Unicode.GetString(body.Slice(cursor, byteCount))
                : Encoding.Latin1.GetString(body.Slice(cursor, byteCount));
            cursor += byteCount;
        }
        else
        {
            int byteCount = Math.Min(nameLength, Math.Max(0, body.Length - cursor));
            name = globals.Encoding.GetString(body.Slice(cursor, byteCount));
            cursor += byteCount;
        }

        if (isBuiltIn && name.Length == 1)
            name = BuiltInName(name[0]);

        rgceLength = Math.Min(rgceLength, Math.Max(0, body.Length - cursor));
        byte[] rgce = body.Slice(cursor, rgceLength).ToArray();
        globals.Names.Add(new XlsDefinedName(name, rgce, sheetIndex, isHidden, isBuiltIn));
    }

    /// <summary>
    /// Expands the single-byte code of one of Excel's own names (MS-XLS 2.5.97.10). These carry
    /// the print area, the print titles and the range an auto filter covers.
    /// </summary>
    private static string BuiltInName(char code) => (int)code switch
    {
        0x00 => "_xlnm.Consolidate_Area",
        0x01 => "_xlnm.Auto_Open",
        0x02 => "_xlnm.Auto_Close",
        0x03 => "_xlnm.Extract",
        0x04 => "_xlnm.Database",
        0x05 => "_xlnm.Criteria",
        0x06 => "_xlnm.Print_Area",
        0x07 => "_xlnm.Print_Titles",
        0x08 => "_xlnm.Recorder",
        0x09 => "_xlnm.Data_Form",
        0x0A => "_xlnm.Auto_Activate",
        0x0B => "_xlnm.Auto_Deactivate",
        0x0C => "_xlnm.Sheet_Title",
        0x0D => "_xlnm._FilterDatabase",
        _ => "_xlnm.Unknown",
    };

    private static XlsSupBook ReadSupBook(ReadOnlySpan<byte> body)
    {
        if (body.Length < 4)
            return new XlsSupBook();

        // A two-byte marker after the sheet count says the record stands for this workbook or
        // for the add-ins rather than for another file.
        ushort marker = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        if (marker == 0x0401)
            return new XlsSupBook();
        if (marker == 0x3A01)
            return new XlsSupBook { IsAddin = true };

        int cursor = 2;
        string path = DecodeVirtualPath(BiffStrings.ReadUnicodeString(body, ref cursor));

        // The other workbook's own tab names follow its path, and its 3-D references count
        // sheets in this list rather than in ours.
        int count = BinaryPrimitives.ReadUInt16LittleEndian(body);
        var names = new List<string>(Math.Min(count, 64));
        for (int i = 0; i < count && cursor < body.Length; i++)
            names.Add(BiffStrings.ReadUnicodeString(body, ref cursor));

        return new XlsSupBook { Path = path, SheetNames = names };
    }

    /// <summary>
    /// Reads an ExternName record, which names one thing borrowed from the SupBook it follows.
    /// For the add-in book these are function names — the only place a file states that a
    /// formula calls EOMONTH rather than a built-in.
    /// </summary>
    private static void ReadExternName(XlsGlobals globals, ReadOnlySpan<byte> body)
    {
        if (globals.SupBooks.Count == 0 || body.Length < 7)
            return;

        int cursor = 6;                          // Flags, an index, and two reserved bytes.
        globals.SupBooks[^1].Names.Add(BiffStrings.ReadShortUnicodeString(body, ref cursor));
    }

    /// <summary>
    /// Turns the control characters of an encoded workbook path into the separators they stand
    /// for (MS-XLS 2.5.277): a path is stored with its volume and directory steps as single
    /// characters below 0x20 rather than spelled out.
    /// </summary>
    private static string DecodeVirtualPath(string encoded)
    {
        var text = new StringBuilder(encoded.Length + 4);
        foreach (char character in encoded)
        {
            switch (character)
            {
                case '\u0001': break;            // A volume follows, given as its letter.
                case '\u0002': text.Append('\\'); break;   // The same volume, from its root.
                case '\u0003': text.Append('\\'); break;
                case '\u0004': text.Append(".."); text.Append('\\'); break;
                case '\u0005': break;            // A long volume name follows.
                case '\u0006': case '\u0007': case '\u0008': break;   // Relative to a startup folder.
                default:
                    if (character >= ' ')
                        text.Append(character);
                    break;
            }
        }

        return text.ToString();
    }

    private static void ReadExternSheet(XlsGlobals globals, ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        int count = BinaryPrimitives.ReadUInt16LittleEndian(body);
        for (int i = 0; i < count && 2 + (i * 6) + 6 <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(2 + (i * 6), 6);
            ushort supBook = BinaryPrimitives.ReadUInt16LittleEndian(entry);
            short first = BinaryPrimitives.ReadInt16LittleEndian(entry[2..]);
            short last = BinaryPrimitives.ReadInt16LittleEndian(entry[4..]);
            globals.ExternSheets.Add((supBook, first, last));
        }
    }

    /// <summary>
    /// Resolves the code page a file declares. 1200 means the strings are already UTF-16, so
    /// the byte-string path never runs; anything the platform cannot provide falls back to
    /// Windows-1252, which is what the great majority of these files actually hold.
    /// </summary>
    private Encoding ResolveEncoding(ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return Encoding.Latin1;

        int codePage = BinaryPrimitives.ReadUInt16LittleEndian(body);
        if (codePage is 1200 or 0)
            return Encoding.Latin1;

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException)
        {
            // Most single-byte code pages need CodePagesEncodingProvider registered, which is
            // the caller's decision to make; without it, Windows-1252 is the closest thing.
            warn(new WorkbookWarning(
                "UnknownCodePage",
                $"The workbook declares code page {codePage}, which is not available here; text was read as " +
                "Windows-1252. Registering CodePagesEncodingProvider before loading widens what can be decoded."));
            return Encoding.Latin1;
        }
    }
}
