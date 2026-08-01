using System.Buffers.Binary;
using System.Text;
using Cellwright.Styles;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads a DXFN block — the formatting a conditional formatting rule paints over a cell that
/// matches. It is an override rather than a style: every property carries a companion bit
/// saying whether it means anything, and the ones that do not are left to the cell's own
/// format.
/// </summary>
internal static class DifferentialFormat
{
    /// <summary>Size of the font block, whatever of it is filled in.</summary>
    private const int FontBlockSize = 118;

    /// <summary>Reads the block at <paramref name="offset"/> and advances past it.</summary>
    /// <param name="body">Record body.</param>
    /// <param name="offset">Where the block starts; advanced to just past it.</param>
    /// <param name="globals">Workbook globals, for resolving palette indexes.</param>
    public static DifferentialStyle Read(ReadOnlySpan<byte> body, ref int offset, XlsGlobals? globals)
    {
        if (offset + 6 > body.Length)
        {
            offset = body.Length;
            return DifferentialStyle.Empty;
        }

        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        ushort extra = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 4)..]);
        offset += 6;

        bool hasNumberFormat = (flags & 0x02000000) != 0;
        bool hasFont = (flags & 0x04000000) != 0;
        bool hasAlignment = (flags & 0x08000000) != 0;
        bool hasBorder = (flags & 0x10000000) != 0;
        bool hasPattern = (flags & 0x20000000) != 0;
        bool hasProtection = (flags & 0x40000000) != 0;
        bool formatIsString = (extra & 0x0001) != 0;

        string? numberFormat = hasNumberFormat ? ReadNumberFormat(body, ref offset, formatIsString) : null;

        bool? bold = null;
        bool? italic = null;
        bool? underline = null;
        bool? strikethrough = null;
        ExcelColor fontColor = ExcelColor.Auto;
        if (hasFont)
            ReadFont(body, ref offset, globals, out bold, out italic, out underline, out strikethrough, out fontColor);

        if (hasAlignment)
            offset += 8;

        BorderStyle? border = hasBorder ? ReadBorder(body, ref offset, flags, globals) : null;
        ExcelColor fill = hasPattern ? ReadFill(body, ref offset, flags, globals) : ExcelColor.Auto;

        if (hasProtection)
            offset += 2;

        offset = Math.Min(offset, body.Length);
        return new DifferentialStyle
        {
            Bold = bold,
            Italic = italic,
            Underline = underline,
            Strikethrough = strikethrough,
            FontColor = fontColor,
            FillColor = fill,
            NumberFormat = numberFormat,
            Border = border,
        };
    }

    private static string? ReadNumberFormat(ReadOnlySpan<byte> body, ref int offset, bool isString)
    {
        if (!isString)
        {
            // A builtin: two bytes, of which the second is the format id.
            if (offset + 2 > body.Length)
                return null;

            int id = body[offset + 1];
            offset += 2;
            return BuiltinNumberFormats.GetCode(id);
        }

        if (offset + 2 > body.Length)
            return null;

        // A user format states the size of the whole block, then the string inside it.
        int size = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        int end = Math.Min(offset + Math.Max(size, 2), body.Length);
        int cursor = offset + 2;
        string code = BiffStrings.ReadUnicodeString(body, ref cursor);
        offset = end;
        return code.Length == 0 ? null : code;
    }

    private static void ReadFont(
        ReadOnlySpan<byte> body,
        ref int offset,
        XlsGlobals? globals,
        out bool? bold,
        out bool? italic,
        out bool? underline,
        out bool? strikethrough,
        out ExcelColor color)
    {
        bold = null;
        italic = null;
        underline = null;
        strikethrough = null;
        color = ExcelColor.Auto;

        if (offset + FontBlockSize > body.Length)
        {
            offset = body.Length;
            return;
        }

        ReadOnlySpan<byte> block = body.Slice(offset, FontBlockSize);
        offset += FontBlockSize;

        // The name and its padding fill the first 64 bytes; the attributes follow.
        uint style = BinaryPrimitives.ReadUInt32LittleEndian(block[68..]);      // stxp.ts
        int weight = BinaryPrimitives.ReadInt16LittleEndian(block[72..]);       // stxp.bls
        byte underlineStyle = block[76];                                        // stxp.uls
        int fontColor = BinaryPrimitives.ReadInt32LittleEndian(block[80..]);    // icvFore
        uint styleNinch = BinaryPrimitives.ReadUInt32LittleEndian(block[88..]); // tsNinch
        bool underlineNinch = BinaryPrimitives.ReadUInt32LittleEndian(block[96..]) != 0;
        bool weightNinch = BinaryPrimitives.ReadUInt32LittleEndian(block[100..]) != 0;

        if (!weightNinch)
            bold = weight >= 700;
        if ((styleNinch & 0x00000002) == 0)
            italic = (style & 0x00000002) != 0;
        if ((styleNinch & 0x00000080) == 0)
            strikethrough = (style & 0x00000080) != 0;
        if (!underlineNinch)
            underline = underlineStyle != 0;

        // -1 means the rule leaves the text color alone; 32767 is "automatic".
        if (fontColor is not (-1 or 0x7FFF))
            color = Resolve(globals, fontColor);
    }

    private static BorderStyle? ReadBorder(
        ReadOnlySpan<byte> body, ref int offset, uint flags, XlsGlobals? globals)
    {
        if (offset + 8 > body.Length)
        {
            offset = body.Length;
            return null;
        }

        uint low = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        uint high = BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 4)..]);
        offset += 8;

        BorderEdge left = Edge((flags & 0x00000400) == 0, (int)(low & 0x0F), (int)((low >> 16) & 0x7F), globals);
        BorderEdge right = Edge((flags & 0x00000800) == 0, (int)((low >> 4) & 0x0F), (int)((low >> 23) & 0x7F), globals);
        BorderEdge top = Edge((flags & 0x00001000) == 0, (int)((low >> 8) & 0x0F), (int)(high & 0x7F), globals);
        BorderEdge bottom = Edge((flags & 0x00002000) == 0, (int)((low >> 12) & 0x0F), (int)((high >> 7) & 0x7F), globals);

        var border = new BorderStyle { Left = left, Right = right, Top = top, Bottom = bottom };
        return border == BorderStyle.None ? null : border;
    }

    private static BorderEdge Edge(bool applies, int lineStyle, int colorIndex, XlsGlobals? globals) =>
        applies
            ? new BorderEdge(MapLine(lineStyle), Resolve(globals, colorIndex))
            : new BorderEdge(BorderLineStyle.None, ExcelColor.Auto);

    /// <summary>
    /// Reads the fill. A rule that paints a cell uses a solid pattern, and it is the pattern's
    /// foreground that shows; the model keeps that one color and nothing else.
    /// </summary>
    private static ExcelColor ReadFill(ReadOnlySpan<byte> body, ref int offset, uint flags, XlsGlobals? globals)
    {
        if (offset + 4 > body.Length)
        {
            offset = body.Length;
            return ExcelColor.Auto;
        }

        ushort patternField = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        ushort colorField = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]);
        offset += 4;

        int pattern = (patternField >> 10) & 0x3F;
        if (pattern == 0)
            return ExcelColor.Auto;

        bool foregroundApplies = (flags & 0x00020000) == 0;
        bool backgroundApplies = (flags & 0x00040000) == 0;
        int foreground = colorField & 0x7F;
        int background = (colorField >> 7) & 0x7F;

        if (foregroundApplies && foreground is not (0 or 64 or 65 or 0x7F))
            return Resolve(globals, foreground);
        if (backgroundApplies && background is not (0 or 64 or 65 or 0x7F))
            return Resolve(globals, background);

        return ExcelColor.Auto;
    }

    private static ExcelColor Resolve(XlsGlobals? globals, int index) =>
        globals?.ColorOf(index) ?? ExcelColor.FromIndexed(index);

    private static BorderLineStyle MapLine(int value) => value switch
    {
        1 => BorderLineStyle.Thin,
        2 => BorderLineStyle.Medium,
        3 => BorderLineStyle.Dashed,
        4 => BorderLineStyle.Dotted,
        5 => BorderLineStyle.Thick,
        6 => BorderLineStyle.Double,
        7 => BorderLineStyle.Hair,
        8 => BorderLineStyle.MediumDashed,
        9 => BorderLineStyle.DashDot,
        10 => BorderLineStyle.MediumDashDot,
        11 => BorderLineStyle.DashDotDot,
        12 => BorderLineStyle.MediumDashDotDot,
        13 => BorderLineStyle.SlantDashDot,
        _ => BorderLineStyle.None,
    };
}
