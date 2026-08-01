using System.Buffers.Binary;
using System.Text;
using Cellwright.Model;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads the HLINK record. The legacy format stores a hyperlink as an OLE moniker — the same
/// structure a shortcut file uses — so the target has to be unpicked from it rather than read
/// out of an attribute (MS-OSHARED 2.3.7).
/// </summary>
internal static class HyperlinkRecord
{
    private static ReadOnlySpan<byte> UrlMoniker =>
        [0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B];

    private static ReadOnlySpan<byte> FileMoniker =>
        [0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];

    private static ReadOnlySpan<byte> CompositeMoniker =>
        [0x09, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];

    private static ReadOnlySpan<byte> AntiMoniker =>
        [0x05, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];

    private static ReadOnlySpan<byte> ItemMoniker =>
        [0x04, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];

    /// <summary>Reads one hyperlink into the sink, ignoring one it cannot make sense of.</summary>
    /// <remarks>
    /// Every optional field of the structure is present only when its own flag says so, and
    /// they follow one another with nothing to resynchronise on. Getting one flag wrong does
    /// not lose one field: it shifts every field after it, so the target of the link ends up
    /// read out of the middle of the display text.
    /// </remarks>
    public static void Read(ReadOnlySpan<byte> body, WorksheetContentSink sink, Action<WorkbookWarning> warn)
    {
        if (body.Length < 32)
            return;

        RangeRef range = ReadRange(body);

        // MS-OSHARED 2.3.7.1. Bit 2 says the author chose the display text rather than letting
        // it default to the target; bit 4 is the one that says the text is actually here.
        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body[28..]);
        bool hasMoniker = (flags & 0x0001) != 0;
        bool isAbsolute = (flags & 0x0002) != 0;
        bool hasLocation = (flags & 0x0008) != 0;
        bool hasDisplay = (flags & 0x0010) != 0;
        bool hasGuid = (flags & 0x0020) != 0;
        bool hasCreationTime = (flags & 0x0040) != 0;
        bool hasFrameName = (flags & 0x0080) != 0;
        bool monikerIsString = (flags & 0x0100) != 0;

        int offset = 32;
        string? display = hasDisplay ? ReadHyperlinkString(body, ref offset) : null;
        if (hasFrameName)
            ReadHyperlinkString(body, ref offset);

        string? target = null;
        if (hasMoniker)
        {
            target = monikerIsString
                ? ReadHyperlinkString(body, ref offset)
                : ReadMoniker(body, ref offset, isAbsolute, warn);
        }

        string? location = hasLocation ? ReadHyperlinkString(body, ref offset) : null;
        _ = hasGuid;
        _ = hasCreationTime;

        sink.OnResolvedHyperlink(range, target, location, tooltip: null, display);
    }

    /// <summary>The cell range an HLINK or HLinkTooltip record applies to (a <c>Ref8U</c>).</summary>
    public static RangeRef ReadRange(ReadOnlySpan<byte> body) => new(
        new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
            BinaryPrimitives.ReadUInt16LittleEndian(body[4..]) + 1),
        new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1,
            BinaryPrimitives.ReadUInt16LittleEndian(body[6..]) + 1));

    private static string? ReadMoniker(
        ReadOnlySpan<byte> body, ref int offset, bool isAbsolute, Action<WorkbookWarning> warn)
    {
        if (offset + 16 > body.Length)
            return null;

        ReadOnlySpan<byte> guid = body.Slice(offset, 16);
        offset += 16;

        if (guid.SequenceEqual(UrlMoniker))
            return ReadUrlMoniker(body, ref offset);
        if (guid.SequenceEqual(FileMoniker))
            return ReadFileMoniker(body, ref offset, isAbsolute);
        if (guid.SequenceEqual(CompositeMoniker))
            return ReadCompositeMoniker(body, ref offset, isAbsolute, warn);
        if (guid.SequenceEqual(AntiMoniker))
            return ReadAntiMoniker(body, ref offset);
        if (guid.SequenceEqual(ItemMoniker))
            return ReadItemMoniker(body, ref offset);

        warn(new WorkbookWarning(
            "UnknownHyperlinkTarget",
            "A hyperlink points at something described by a moniker this library does not read; only its text survived."));
        return null;
    }

    /// <summary>
    /// A URL moniker (MS-OSHARED 2.3.7.6). Its length covers the URL and, when it is 24 bytes
    /// longer, a trailing serialisation block — so the text ends at the first NUL rather than
    /// at the end of what the length counts.
    /// </summary>
    private static string? ReadUrlMoniker(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return null;

        int length = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        length = Math.Min(length, Math.Max(0, body.Length - offset));

        string url = ReadNulTerminated(body.Slice(offset, length));
        offset += length;
        return url;
    }

    /// <summary>
    /// A composite moniker (MS-OSHARED 2.3.7.3): several monikers in sequence, which is how a
    /// path and the item within it are written together. Their texts concatenate.
    /// </summary>
    private static string? ReadCompositeMoniker(
        ReadOnlySpan<byte> body, ref int offset, bool isAbsolute, Action<WorkbookWarning> warn)
    {
        if (offset + 4 > body.Length)
            return null;

        int count = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (count is <= 0 or > 64)
            return null;

        var parts = new List<string>(count);
        for (int i = 0; i < count && offset < body.Length; i++)
        {
            if (ReadMoniker(body, ref offset, isAbsolute, warn) is { Length: > 0 } part)
                parts.Add(part);
        }

        return parts.Count == 0 ? null : string.Concat(parts);
    }

    /// <summary>
    /// An anti-moniker (MS-OSHARED 2.3.7.4) cancels the moniker before it in a composite, which
    /// as text is one step up the directory tree per count.
    /// </summary>
    private static string? ReadAntiMoniker(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return null;

        int count = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        return count is > 0 and <= 1_048_576 ? string.Concat(Enumerable.Repeat("../", count)) : null;
    }

    /// <summary>
    /// An item moniker (MS-OSHARED 2.3.7.5): a delimiter and the name of something inside the
    /// container the preceding moniker named, such as a range within a workbook.
    /// </summary>
    private static string? ReadItemMoniker(ReadOnlySpan<byte> body, ref int offset)
    {
        string? delimiter = ReadItemString(body, ref offset);
        string? item = ReadItemString(body, ref offset);
        if (item is null)
            return null;

        return delimiter is { Length: > 0 } ? delimiter + item : item;
    }

    /// <summary>
    /// One counted string of an item moniker: ANSI characters, and Unicode ones after them when
    /// the count says the ANSI form did not hold everything.
    /// </summary>
    private static string? ReadItemString(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return null;

        int length = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (length < 0 || offset + length > body.Length)
            return null;

        ReadOnlySpan<byte> counted = body.Slice(offset, length);
        offset += length;

        int nul = counted.IndexOf((byte)0);
        if (nul < 0)
            return Encoding.Latin1.GetString(counted);

        string ansi = Encoding.Latin1.GetString(counted[..nul]);
        ReadOnlySpan<byte> wide = counted[(nul + 1)..];
        return wide.Length >= 2 ? ReadNulTerminated(wide) : ansi;
    }

    private static string? ReadFileMoniker(ReadOnlySpan<byte> body, ref int offset, bool isAbsolute)
    {
        if (offset + 6 > body.Length)
            return null;

        int upLevels = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        int ansiLength = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 2)..]);
        offset += 6;

        ansiLength = Math.Min(ansiLength, Math.Max(0, body.Length - offset));
        string shortPath = Encoding.Latin1.GetString(body.Slice(offset, ansiLength)).TrimEnd('\0');
        offset += ansiLength;

        // Then endServer, versionNumber and two reserved fields — 24 bytes in all — before the
        // optional block that carries the path as Unicode, which is the one worth keeping.
        string path = shortPath;
        if (offset + 24 <= body.Length)
        {
            offset += 24;
            if (offset + 4 <= body.Length)
            {
                int longSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
                offset += 4;
                if (longSize is > 6 and < 0x00100000 && offset + longSize <= body.Length)
                {
                    // The block restates its own length and a constant before the characters,
                    // and unlike everywhere else here they are not null-terminated.
                    int characters = Math.Max(0, longSize - 6);
                    path = Encoding.Unicode.GetString(body.Slice(offset + 6, characters)).TrimEnd('\0');
                    offset += longSize;
                }
            }
        }

        if (path.Length == 0)
            return null;

        return isAbsolute || upLevels == 0 ? path : string.Concat(Enumerable.Repeat("../", upLevels)) + path;
    }

    /// <summary>
    /// Reads a <c>HyperlinkString</c> (MS-OSHARED 2.3.7.9): a count of UTF-16 characters that
    /// includes the terminating NUL, then the characters.
    /// </summary>
    private static string ReadHyperlinkString(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return "";

        int characters = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        int byteCount = Math.Min(Math.Max(characters, 0) * 2, Math.Max(0, body.Length - offset - 4));
        offset += 4;
        string text = ReadNulTerminated(body.Slice(offset, byteCount));
        offset += byteCount;
        return text;
    }

    /// <summary>Decodes UTF-16 up to the first NUL, ignoring whatever follows it.</summary>
    private static string ReadNulTerminated(ReadOnlySpan<byte> utf16)
    {
        for (int i = 0; i + 1 < utf16.Length; i += 2)
        {
            if (utf16[i] == 0 && utf16[i + 1] == 0)
                return Encoding.Unicode.GetString(utf16[..i]);
        }

        return Encoding.Unicode.GetString(utf16[..(utf16.Length & ~1)]);
    }
}

/// <summary>
/// Reads the HLinkTooltip record (MS-XLS 2.4.141), which follows the HLink it belongs to and
/// carries the text shown when the pointer rests on the link. Excel keeps the tooltip here and
/// not in the hyperlink structure itself, so a reader that looks only at HLink never finds one.
/// </summary>
internal static class HyperlinkTooltipRecord
{
    /// <summary>Attaches the tooltip to the hyperlink over the same range.</summary>
    public static void Read(ReadOnlySpan<byte> body, Worksheet sheet)
    {
        // FrtRefHeaderNoGrbit: the record type again, then the range, then the text.
        if (body.Length < 12)
            return;

        RangeRef range = HyperlinkRecord.ReadRange(body[2..]);
        string tooltip = ReadNulTerminated(body[10..]);
        if (tooltip.Length == 0)
            return;

        foreach (Hyperlink hyperlink in sheet.Hyperlinks)
        {
            if (hyperlink.Range == range)
            {
                hyperlink.Tooltip = tooltip;
                return;
            }
        }
    }

    private static string ReadNulTerminated(ReadOnlySpan<byte> utf16)
    {
        for (int i = 0; i + 1 < utf16.Length; i += 2)
        {
            if (utf16[i] == 0 && utf16[i + 1] == 0)
                return Encoding.Unicode.GetString(utf16[..i]);
        }

        return Encoding.Unicode.GetString(utf16[..(utf16.Length & ~1)]);
    }
}

/// <summary>Reads the DV record, one input rule and the messages that go with it.</summary>
internal static class DataValidationRecord
{
    /// <summary>Reads one rule into the sink.</summary>
    public static void Read(
        ReadOnlySpan<byte> body, WorksheetContentSink sink, PtgDecoder decoder, Action<WorkbookWarning> warn)
    {
        if (body.Length < 4)
            return;

        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body);
        int offset = 4;
        string promptTitle = ReadUnicodeString(body, ref offset);
        string errorTitle = ReadUnicodeString(body, ref offset);
        string prompt = ReadUnicodeString(body, ref offset);
        string error = ReadUnicodeString(body, ref offset);

        string? formula1 = ReadFormula(body, ref offset, decoder);
        string? formula2 = ReadFormula(body, ref offset, decoder);
        IReadOnlyList<RangeRef> ranges = ReadRanges(body, ref offset);
        if (ranges.Count == 0)
            return;

        DataValidationType type = MapType((int)(flags & 0x0F));
        if (type == DataValidationType.List)
            formula1 = UnpackLiteralList(formula1);

        var validation = new DataValidation
        {
            Ranges = ranges,
            Type = type,
            Operator = MapOperator((int)((flags >> 20) & 0x0F)),
            Formula1 = formula1,
            Formula2 = formula2,
            AllowBlank = (flags & 0x00000100) != 0,
            ShowDropDown = (flags & 0x00000200) == 0,
            ShowInputMessage = (flags & 0x00040000) != 0,
            ShowErrorMessage = (flags & 0x00080000) != 0,
            ErrorStyle = MapErrorStyle((int)((flags >> 4) & 0x07)),
            PromptTitle = Nullify(promptTitle),
            Prompt = Nullify(prompt),
            ErrorTitle = Nullify(errorTitle),
            Error = Nullify(error),
        };

        if (validation.Type == DataValidationType.Custom && formula1 is null)
        {
            warn(new WorkbookWarning(
                "UndecodedValidation",
                $"A custom validation rule on '{sink.Sheet.Name}' uses a formula this library cannot render and was dropped."));
            return;
        }

        sink.OnDataValidation(validation);
    }

    private static string? ReadFormula(ReadOnlySpan<byte> body, ref int offset, PtgDecoder decoder)
    {
        if (offset + 4 > body.Length)
            return null;

        int length = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += 4; // The two bytes after the length are reserved.
        length = Math.Min(length, Math.Max(0, body.Length - offset));
        if (length == 0)
            return null;

        // A rule's formula is anchored at A1: it has no cell of its own to be relative to.
        string? text = decoder.Decode(body.Slice(offset, length), new CellRef(1, 1), sheetIndex: 0);
        offset += length;
        return text;
    }

    private static IReadOnlyList<RangeRef> ReadRanges(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 2 > body.Length)
            return [];

        int count = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += 2;

        var ranges = new List<RangeRef>(count);
        for (int i = 0; i < count && offset + 8 <= body.Length; i++)
        {
            ranges.Add(new RangeRef(
                new CellRef(
                    BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]) + 1,
                    BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 4)..]) + 1),
                new CellRef(
                    BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]) + 1,
                    BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 6)..]) + 1)));
            offset += 8;
        }

        return ranges;
    }

    private static string ReadUnicodeString(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 3 > body.Length)
        {
            offset = body.Length;
            return "";
        }

        int characters = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        bool wide = (body[offset + 2] & 0x01) != 0;
        offset += 3;

        int byteCount = Math.Min(wide ? characters * 2 : characters, Math.Max(0, body.Length - offset));
        string text = wide
            ? Encoding.Unicode.GetString(body.Slice(offset, byteCount))
            : Encoding.Latin1.GetString(body.Slice(offset, byteCount));
        offset += byteCount;
        return text;
    }

    /// <summary>A single NUL is how the format spells "no message here".</summary>
    private static string? Nullify(string text) => text.Length == 0 || text == "\0" ? null : text;

    /// <summary>
    /// Turns a list of literal choices back into the comma-separated string the model and the
    /// modern format use. The binary format stores them as one string with NUL between them.
    /// </summary>
    private static string? UnpackLiteralList(string? formula) =>
        formula is not null && formula.Contains('\0', StringComparison.Ordinal)
            ? formula.Replace('\0', ',')
            : formula;

    /// <summary>
    /// Turns a list of literal choices back into the comma-separated string the model uses,
    /// where the binary formats put a NUL between them. Shared with the BIFF12 reader.
    /// </summary>
    public static string? UnpackList(string? formula) => UnpackLiteralList(formula);

    /// <summary>
    /// The three enumerations a validation rule is built from. BIFF12 numbers them exactly as
    /// BIFF8 does (MS-XLSB 2.4.356 against MS-XLS 2.4.90), so both readers share these rather
    /// than keeping two copies of a table that has to agree.
    /// </summary>
    public static DataValidationType MapType(int value) => value switch
    {
        1 => DataValidationType.Whole,
        2 => DataValidationType.Decimal,
        3 => DataValidationType.List,
        4 => DataValidationType.Date,
        5 => DataValidationType.Time,
        6 => DataValidationType.TextLength,
        7 => DataValidationType.Custom,
        _ => DataValidationType.None,
    };

    /// <inheritdoc cref="MapType"/>
    public static DataValidationOperator MapOperator(int value) => value switch
    {
        1 => DataValidationOperator.NotBetween,
        2 => DataValidationOperator.Equal,
        3 => DataValidationOperator.NotEqual,
        4 => DataValidationOperator.GreaterThan,
        5 => DataValidationOperator.LessThan,
        6 => DataValidationOperator.GreaterThanOrEqual,
        7 => DataValidationOperator.LessThanOrEqual,
        _ => DataValidationOperator.Between,
    };

    /// <inheritdoc cref="MapType"/>
    public static DataValidationErrorStyle MapErrorStyle(int value) => value switch
    {
        1 => DataValidationErrorStyle.Warning,
        2 => DataValidationErrorStyle.Information,
        _ => DataValidationErrorStyle.Stop,
    };
}
