using Cellwright.Styles;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Turns the workbook's style pool into the four tables the file format keeps: fonts, number
/// formats, the palette, and the XF records that combine them.
/// </summary>
/// <remarks>
/// The pool holds whole styles; the file holds indexes into shared sub-tables. Building it is
/// therefore a deduplication pass over every style in the workbook, done once before anything
/// is written because a cell names its XF by position.
/// </remarks>
internal sealed class StyleTable
{
    /// <summary>Excel expects fifteen style XFs before the first cell one (MS-XLS 2.4.353).</summary>
    private const int StyleXfCount = 15;

    /// <summary>The palette has 56 slots, numbered from 8.</summary>
    private const int PaletteSize = 56;

    private const int FirstPaletteIndex = 8;

    /// <summary>The format ids a workbook may define for itself.</summary>
    private const int FirstCustomFormatId = 164;

    private readonly List<FontStyle> _fonts = [];
    private readonly List<(int Id, string Code)> _formats = [];
    private readonly Dictionary<string, int> _formatIds = new(StringComparer.Ordinal);
    private readonly List<XfEntry> _cellFormats = [];
    private readonly Dictionary<int, int> _xfByStyle = [];
    private readonly List<TemporalKind> _temporalByXf = [];
    private readonly List<uint> _palette = [];
    private readonly Dictionary<uint, int> _paletteIndexes = [];

    private StyleTable()
    {
    }

    /// <summary>The XF index each style-pool entry maps to.</summary>
    public IReadOnlyDictionary<int, int> XfByStyle => _xfByStyle;

    /// <summary>Whether each XF formats its number as a date or a duration.</summary>
    public IReadOnlyList<TemporalKind> TemporalByXf => _temporalByXf;

    /// <summary>The colour palette, or empty when the defaults suffice.</summary>
    public IReadOnlyList<uint> Palette => _palette;

    /// <summary>Builds the tables from every style the workbook has interned.</summary>
    public static StyleTable Build(StylePool pool, Action<WorkbookWarning> warn)
    {
        var table = new StyleTable();

        // The default font is always index 0, and index 4 is skipped: the format reserves it
        // for reasons long since lost, and Excel expects the gap.
        table._fonts.Add(FontStyle.Default);

        // The style XFs come first and are not addressable by a cell, so they are filled with
        // the default and never looked up again.
        for (int i = 0; i < StyleXfCount; i++)
        {
            table._cellFormats.Add(XfEntry.Default with { IsStyleXf = true });
            table._temporalByXf.Add(TemporalKind.None);
        }

        for (int index = 0; index < pool.Styles.Count; index++)
        {
            CellStyle style = pool.Styles[index];
            table._xfByStyle[index] = table.AddStyle(style);
        }

        if (table._cellFormats.Count > 4050)
        {
            warn(new WorkbookWarning(
                "TooManyStyles",
                $"The workbook has {table._cellFormats.Count} distinct cell formats; the .xls format holds 4050, " +
                "so the ones past that limit fall back to the default format."));
        }

        return table;
    }

    private int AddStyle(CellStyle style)
    {
        int fontIndex = InternFont(style.Font);
        int formatId = InternFormat(style.NumberFormat);

        var entry = new XfEntry
        {
            FontIndex = fontIndex,
            FormatId = formatId,
            Style = style,
            ForegroundColor = InternColor(style.Fill.Color),
            BackgroundColor = InternColor(style.Fill.BackgroundColor),
            LeftColor = InternColor(style.Border.Left.Color),
            RightColor = InternColor(style.Border.Right.Color),
            TopColor = InternColor(style.Border.Top.Color),
            BottomColor = InternColor(style.Border.Bottom.Color),
            DiagonalColor = InternColor(style.Border.Diagonal.Color),
        };

        // Past the format's own limit a cell has to point somewhere; the default is the only
        // index guaranteed to exist.
        if (_cellFormats.Count >= 4050)
            return XlsWriteContext.DefaultCellXf;

        _cellFormats.Add(entry);
        _temporalByXf.Add(NumberFormatClassifier.Classify(
            formatId, style.NumberFormat.Length == 0 ? null : style.NumberFormat));
        return _cellFormats.Count - 1;
    }

    private int InternFont(FontStyle font)
    {
        for (int i = 0; i < _fonts.Count; i++)
        {
            if (_fonts[i] == font)
                return FontSlot(i);
        }

        _fonts.Add(font);
        return FontSlot(_fonts.Count - 1);
    }

    /// <summary>Maps a position in the font list onto its index in the file, skipping slot 4.</summary>
    private static int FontSlot(int position) => position < 4 ? position : position + 1;

    private int InternFormat(string code)
    {
        if (code.Length == 0)
            return 0;

        if (BuiltinNumberFormats.GetId(code) is { } builtin)
            return builtin;

        if (_formatIds.TryGetValue(code, out int id))
            return id;

        id = FirstCustomFormatId + _formats.Count;
        _formatIds[code] = id;
        _formats.Add((id, code));
        return id;
    }

    /// <summary>
    /// Places a colour in the palette, reusing a slot whose colour is already there. Past 56
    /// distinct colours the nearest existing slot is used, since the format has nowhere else to
    /// put one.
    /// </summary>
    /// <summary>The palette slot a colour occupies, adding it when the palette has room.</summary>
    /// <param name="color">The colour to place.</param>
    /// <remarks>
    /// Public because a conditional formatting rule paints in colours no cell style names, and
    /// those have to reach the palette too. Safe to call while the sheets are being written: the
    /// palette is emitted with the globals, which BIFF puts after them.
    /// </remarks>
    public int ColorIndex(ExcelColor color) => InternColor(color);

    private int InternColor(ExcelColor color)
    {
        // The theme is an xlsx idea; the old format has no way to express one, so a themed
        // colour falls back to automatic rather than to a wrong literal.
        if (color.IsAuto || color.IsThemed)
            return 0x7FFF;

        uint argb = color.Argb | 0xFF000000u;
        if (_paletteIndexes.TryGetValue(argb, out int index))
            return index;

        if (_palette.Count < PaletteSize)
        {
            index = FirstPaletteIndex + _palette.Count;
            _palette.Add(argb);
            _paletteIndexes[argb] = index;
            return index;
        }

        return NearestSlot(argb);
    }

    /// <summary>Finds the palette slot whose colour is closest, by squared distance in RGB.</summary>
    private int NearestSlot(uint argb)
    {
        int red = (int)((argb >> 16) & 0xFF);
        int green = (int)((argb >> 8) & 0xFF);
        int blue = (int)(argb & 0xFF);

        int best = FirstPaletteIndex;
        long bestDistance = long.MaxValue;
        for (int i = 0; i < _palette.Count; i++)
        {
            uint candidate = _palette[i];
            long dr = red - (int)((candidate >> 16) & 0xFF);
            long dg = green - (int)((candidate >> 8) & 0xFF);
            long db = blue - (int)(candidate & 0xFF);
            long distance = (dr * dr) + (dg * dg) + (db * db);
            if (distance >= bestDistance)
                continue;

            bestDistance = distance;
            best = FirstPaletteIndex + i;
        }

        return best;
    }

    /// <summary>Writes the fonts, formats and XF records in the order the format requires.</summary>
    public void Write(BiffWriter writer)
    {
        foreach (FontStyle font in _fonts)
            WriteFont(writer, font);

        foreach ((int id, string code) in _formats)
        {
            int start = writer.BeginRecord(BiffRecordType.Format);
            writer.AppendUInt16(id);
            writer.AppendUnicodeString(code);
            writer.EndRecord(start);
        }

        foreach (XfEntry entry in _cellFormats)
            WriteXf(writer, entry);

        WriteBuiltInStyles(writer);
    }

    private static void WriteFont(BiffWriter writer, FontStyle font)
    {
        int start = writer.BeginRecord(BiffRecordType.Font);
        writer.AppendUInt16((int)Math.Round(font.Size * 20));

        int attributes = 0;
        if (font.Italic)
            attributes |= 0x0002;
        if (font.Strikethrough)
            attributes |= 0x0008;
        writer.AppendUInt16(attributes);

        writer.AppendUInt16(0x7FFF);                              // Automatic colour.
        writer.AppendUInt16(font.Bold ? 700 : 400);
        writer.AppendUInt16(font.VerticalAlign switch
        {
            FontVerticalAlign.Superscript => 1,
            FontVerticalAlign.Subscript => 2,
            _ => 0,
        });
        writer.AppendByte(font.Underline switch
        {
            UnderlineStyle.Single => 0x01,
            UnderlineStyle.Double => 0x02,
            UnderlineStyle.SingleAccounting => 0x21,
            UnderlineStyle.DoubleAccounting => 0x22,
            _ => 0x00,
        });
        writer.AppendByte((byte)Math.Clamp(font.Family, 0, 5));
        writer.AppendByte(font.Charset < 0 ? (byte)0 : (byte)Math.Clamp(font.Charset, 0, 255));
        writer.AppendByte(0);                                     // Reserved.
        writer.AppendShortUnicodeString(font.Name);
        writer.EndRecord(start);
    }

    private static void WriteXf(BiffWriter writer, in XfEntry entry)
    {
        CellStyle style = entry.Style;
        int start = writer.BeginRecord(BiffRecordType.Xf);
        writer.AppendUInt16(entry.FontIndex);
        writer.AppendUInt16(entry.FormatId);

        int protection = 0;
        if (style.Protection.Locked)
            protection |= 0x0001;
        if (style.Protection.Hidden)
            protection |= 0x0002;
        if (entry.IsStyleXf)
            protection |= 0x0004;

        // A cell XF says which style XF it inherits from; 0xFFF means the default one.
        protection |= entry.IsStyleXf ? 0xFFF0 : 0x0000;
        writer.AppendUInt16(protection);

        writer.AppendByte(EncodeAlignment(style.Alignment));
        writer.AppendByte(EncodeRotation(style.Alignment.TextRotation));
        writer.AppendByte(EncodeIndentation(style.Alignment, style.Border));
        writer.AppendByte(0x00);                                  // No attribute is inherited.

        writer.AppendUInt16(EncodeBorderStyles(style.Border));
        writer.AppendUInt16(EncodeBorderColors1(entry, style.Border));
        writer.AppendUInt32(EncodeBorderColors2(entry, style.Fill));
        writer.AppendUInt16(EncodeFill(entry, style.Fill));
        writer.EndRecord(start);
    }

    private static byte EncodeAlignment(Alignment alignment)
    {
        int value = alignment.Horizontal switch
        {
            HorizontalAlignment.Left => 1,
            HorizontalAlignment.Center => 2,
            HorizontalAlignment.Right => 3,
            HorizontalAlignment.Fill => 4,
            HorizontalAlignment.Justify => 5,
            HorizontalAlignment.CenterContinuous => 6,
            HorizontalAlignment.Distributed => 7,
            _ => 0,
        };

        if (alignment.WrapText)
            value |= 0x08;

        value |= alignment.Vertical switch
        {
            VerticalAlignment.Top => 0 << 4,
            VerticalAlignment.Center => 1 << 4,
            VerticalAlignment.Justify => 3 << 4,
            VerticalAlignment.Distributed => 4 << 4,
            _ => 2 << 4,
        };

        return (byte)value;
    }

    private static byte EncodeRotation(int rotation) => rotation switch
    {
        Alignment.StackedRotation => 255,
        >= 0 and <= 180 => (byte)rotation,
        _ => 0,
    };

    private static byte EncodeIndentation(Alignment alignment, BorderStyle border)
    {
        int value = Math.Clamp(alignment.Indent, 0, 15);
        if (alignment.ShrinkToFit)
            value |= 0x10;
        if (border.DiagonalDown)
            value |= 0x20;
        if (border.DiagonalUp)
            value |= 0x40;
        value |= (int)alignment.ReadingOrder << 6;
        return (byte)value;
    }

    private static ushort EncodeBorderStyles(BorderStyle border) => (ushort)(
        MapLine(border.Left.Line) |
        (MapLine(border.Right.Line) << 4) |
        (MapLine(border.Top.Line) << 8) |
        (MapLine(border.Bottom.Line) << 12));

    private static ushort EncodeBorderColors1(in XfEntry entry, BorderStyle border)
    {
        int value = (entry.LeftColor & 0x7F) | ((entry.RightColor & 0x7F) << 7);
        if (border.DiagonalUp)
            value |= 0x4000;
        if (border.DiagonalDown)
            value |= 0x8000;
        return (ushort)value;
    }

    private static uint EncodeBorderColors2(in XfEntry entry, FillStyle fill)
    {
        uint value = (uint)(entry.TopColor & 0x7F);
        value |= (uint)(entry.BottomColor & 0x7F) << 7;
        value |= (uint)(entry.DiagonalColor & 0x7F) << 14;
        value |= (uint)MapLine(entry.Style.Border.Diagonal.Line) << 21;
        value |= (uint)MapPattern(fill.Pattern) << 26;
        return value;
    }

    private static ushort EncodeFill(in XfEntry entry, FillStyle fill)
    {
        // A fill with no explicit colours still needs the pair Excel expects for a solid block:
        // slot 64 is the window foreground and 65 the background.
        int foreground = fill.Pattern == FillPattern.None ? 64 : ClampSlot(entry.ForegroundColor, 64);
        int background = fill.Pattern == FillPattern.None ? 65 : ClampSlot(entry.BackgroundColor, 65);
        return (ushort)((foreground & 0x7F) | ((background & 0x7F) << 7));
    }

    private static int ClampSlot(int index, int fallback) => index == 0x7FFF ? fallback : index;

    private static int MapLine(BorderLineStyle line) => line switch
    {
        BorderLineStyle.Thin => 1,
        BorderLineStyle.Medium => 2,
        BorderLineStyle.Dashed => 3,
        BorderLineStyle.Dotted => 4,
        BorderLineStyle.Thick => 5,
        BorderLineStyle.Double => 6,
        BorderLineStyle.Hair => 7,
        BorderLineStyle.MediumDashed => 8,
        BorderLineStyle.DashDot => 9,
        BorderLineStyle.MediumDashDot => 10,
        BorderLineStyle.DashDotDot => 11,
        BorderLineStyle.MediumDashDotDot => 12,
        BorderLineStyle.SlantDashDot => 13,
        _ => 0,
    };

    private static int MapPattern(FillPattern pattern) => pattern switch
    {
        FillPattern.Solid => 1,
        FillPattern.MediumGray => 2,
        FillPattern.DarkGray => 3,
        FillPattern.LightGray => 4,
        FillPattern.DarkHorizontal => 5,
        FillPattern.DarkVertical => 6,
        FillPattern.DarkDown => 7,
        FillPattern.DarkUp => 8,
        FillPattern.DarkGrid => 9,
        FillPattern.DarkTrellis => 10,
        FillPattern.LightHorizontal => 11,
        FillPattern.LightVertical => 12,
        FillPattern.LightDown => 13,
        FillPattern.LightUp => 14,
        FillPattern.LightGrid => 15,
        FillPattern.LightTrellis => 16,
        FillPattern.Gray125 => 17,
        FillPattern.Gray0625 => 18,
        _ => 0,
    };

    /// <summary>
    /// Writes the STYLE records naming the built-in style XFs. Excel rebuilds its style gallery
    /// from these, and a file without them opens with the gallery empty.
    /// </summary>
    private static void WriteBuiltInStyles(BiffWriter writer)
    {
        // Style 0 is Normal; the rest are the comma, currency and percent variants that make up
        // the fixed fifteen.
        ReadOnlySpan<byte> builtInIds = [0x00, 0x03, 0x04, 0x05, 0x06, 0x07];
        for (int i = 0; i < builtInIds.Length; i++)
        {
            int start = writer.BeginRecord(BiffRecordType.Style);
            writer.AppendUInt16(0x8000 | i);   // Built-in, XF index i.
            writer.AppendByte(builtInIds[i]);
            writer.AppendByte(0xFF);           // No outline level.
            writer.EndRecord(start);
        }
    }

    /// <summary>One row of the XF table, with its colours already placed in the palette.</summary>
    private readonly record struct XfEntry
    {
        public static XfEntry Default { get; } = new()
        {
            Style = CellStyle.Default,
            ForegroundColor = 0x7FFF,
            BackgroundColor = 0x7FFF,
            LeftColor = 0x7FFF,
            RightColor = 0x7FFF,
            TopColor = 0x7FFF,
            BottomColor = 0x7FFF,
            DiagonalColor = 0x7FFF,
        };

        public CellStyle Style { get; init; }

        public int FontIndex { get; init; }

        public int FormatId { get; init; }

        public bool IsStyleXf { get; init; }

        public int ForegroundColor { get; init; }

        public int BackgroundColor { get; init; }

        public int LeftColor { get; init; }

        public int RightColor { get; init; }

        public int TopColor { get; init; }

        public int BottomColor { get; init; }

        public int DiagonalColor { get; init; }
    }
}
