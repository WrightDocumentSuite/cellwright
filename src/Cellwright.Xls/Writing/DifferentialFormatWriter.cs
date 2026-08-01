using Cellwright.Styles;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes a DXFN block — the formatting a conditional formatting rule paints over a cell that
/// matches (MS-XLS 2.5.95).
/// </summary>
/// <remarks>
/// A differential format is an override rather than a style, and the record says so twice: a
/// header flag per block saying whether the block is there at all, and, inside the blocks, a bit
/// per property saying whether that property means anything. Both senses are inverted from what
/// one would expect — a set bit in the header means the block follows, while a set bit inside
/// means "not inherited... changed", so the properties a rule leaves alone are the ones whose
/// bits stay set. Getting that backwards produces a file that reads back as a rule which resets
/// every property it never mentioned.
/// </remarks>
internal static class DifferentialFormatWriter
{
    /// <summary>Size of the font block, whatever of it is filled in.</summary>
    private const int FontBlockSize = 118;

    private const uint HasNumberFormat = 0x02000000;
    private const uint HasFont = 0x04000000;
    private const uint HasBorder = 0x10000000;
    private const uint HasPattern = 0x20000000;

    /// <summary>Set when the corresponding edge or colour is <em>not</em> part of the override.</summary>
    private const uint BorderLeftUnused = 0x00000400;
    private const uint BorderRightUnused = 0x00000800;
    private const uint BorderTopUnused = 0x00001000;
    private const uint BorderBottomUnused = 0x00002000;
    private const uint ForegroundUnused = 0x00020000;
    private const uint BackgroundUnused = 0x00040000;

    /// <summary>The colour index that means "leave the cell's own".</summary>
    private const int NoColor = 0x7FFF;

    /// <summary>Writes the block for a rule's format, or an empty one when it has none.</summary>
    /// <param name="body">Where to write.</param>
    /// <param name="style">The format to write.</param>
    /// <param name="context">The write context, for the palette.</param>
    public static void Write(BiffWriter body, DifferentialStyle? style, XlsWriteContext context)
    {
        if (style is null || style.IsEmpty)
        {
            // A rule that paints nothing still needs a well-formed block saying so.
            body.AppendUInt32(0);
            body.AppendUInt16(0);
            return;
        }

        bool hasFont = style.Bold is not null || style.Italic is not null ||
            style.Underline is not null || style.Strikethrough is not null || !style.FontColor.IsAuto;
        bool hasBorder = style.Border is not null;
        bool hasFill = !style.FillColor.IsAuto;
        bool hasFormat = style.NumberFormat is not null;

        uint flags = 0;
        if (hasFormat)
            flags |= HasNumberFormat;
        if (hasFont)
            flags |= HasFont;
        if (hasBorder)
            flags |= HasBorder;
        if (hasFill)
            flags |= HasPattern;

        if (hasBorder)
            flags |= UnusedEdges(style.Border!);
        else
            flags |= BorderLeftUnused | BorderRightUnused | BorderTopUnused | BorderBottomUnused;

        // Only the pattern's foreground carries the colour a rule paints with; the background is
        // left to the cell, which is how the reader takes it back.
        flags |= BackgroundUnused;
        if (!hasFill)
            flags |= ForegroundUnused;

        body.AppendUInt32(flags);

        // The low bit says the number format is a string rather than a builtin identifier, which
        // is the only form this writes: a code is what the model keeps.
        body.AppendUInt16(hasFormat ? 1 : 0);

        if (hasFormat)
            WriteNumberFormat(body, style.NumberFormat!);
        if (hasFont)
            WriteFont(body, style, context);
        if (hasBorder)
            WriteBorder(body, style.Border!, context);
        if (hasFill)
            WriteFill(body, style.FillColor, context);
    }

    /// <summary>Writes the number format: the size of the block, then the code inside it.</summary>
    private static void WriteNumberFormat(BiffWriter body, string code)
    {
        using var text = new BiffWriter(64);
        text.AppendUnicodeString(code);

        // The stated size counts the size field itself along with the string.
        body.AppendUInt16(text.Length + 2);
        body.Append(text.Written.Span);
    }

    /// <summary>
    /// Writes the 118-byte font block. Most of it is the family name and its padding, which a
    /// rule never overrides, so what is written is a zeroed block with the few fields the model
    /// expresses filled in at the offsets the reader takes them from.
    /// </summary>
    private static void WriteFont(BiffWriter body, DifferentialStyle style, XlsWriteContext context)
    {
        Span<byte> block = stackalloc byte[FontBlockSize];
        block.Clear();

        uint textStyle = 0;
        uint styleNinch = 0;

        // Italic and strikethrough live in one field with one "unchanged" bit each.
        if (style.Italic is bool italic)
        {
            if (italic)
                textStyle |= 0x00000002;
        }
        else
        {
            styleNinch |= 0x00000002;
        }

        if (style.Strikethrough is bool strikethrough)
        {
            if (strikethrough)
                textStyle |= 0x00000080;
        }
        else
        {
            styleNinch |= 0x00000080;
        }

        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(block[68..], textStyle);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(block[88..], styleNinch);

        // Weight is a number rather than a flag: 700 is bold and 400 is not.
        System.Buffers.Binary.BinaryPrimitives.WriteInt16LittleEndian(
            block[72..], style.Bold is bool bold ? (short)(bold ? 700 : 400) : (short)0);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(
            block[100..], style.Bold is null ? 1u : 0u);

        block[76] = style.Underline is true ? (byte)1 : (byte)0;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(
            block[96..], style.Underline is null ? 1u : 0u);

        // -1 is what the reader takes as "the rule leaves the colour alone".
        int color = style.FontColor.IsAuto ? -1 : context.StyleTable.ColorIndex(style.FontColor);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(block[80..], color);

        body.Append(block);
    }

    private static void WriteBorder(BiffWriter body, BorderStyle border, XlsWriteContext context)
    {
        uint low = 0;
        low |= (uint)Line(border.Left.Line) & 0x0F;
        low |= ((uint)Line(border.Right.Line) & 0x0F) << 4;
        low |= ((uint)Line(border.Top.Line) & 0x0F) << 8;
        low |= ((uint)Line(border.Bottom.Line) & 0x0F) << 12;
        low |= ((uint)Color(context, border.Left.Color) & 0x7F) << 16;
        low |= ((uint)Color(context, border.Right.Color) & 0x7F) << 23;

        uint high = (uint)Color(context, border.Top.Color) & 0x7F;
        high |= ((uint)Color(context, border.Bottom.Color) & 0x7F) << 7;

        body.AppendUInt32(low);
        body.AppendUInt32(high);
    }

    /// <summary>The edges a border override says nothing about, as their "unused" bits.</summary>
    private static uint UnusedEdges(BorderStyle border)
    {
        uint flags = 0;
        if (border.Left.Line == BorderLineStyle.None)
            flags |= BorderLeftUnused;
        if (border.Right.Line == BorderLineStyle.None)
            flags |= BorderRightUnused;
        if (border.Top.Line == BorderLineStyle.None)
            flags |= BorderTopUnused;
        if (border.Bottom.Line == BorderLineStyle.None)
            flags |= BorderBottomUnused;

        return flags;
    }

    /// <summary>Writes the fill: a solid pattern whose foreground is the colour the rule paints.</summary>
    private static void WriteFill(BiffWriter body, ExcelColor color, XlsWriteContext context)
    {
        const int SolidPattern = 1;

        body.AppendUInt16(SolidPattern << 10);

        int foreground = Color(context, color);
        body.AppendUInt16((foreground & 0x7F) | ((NoColor & 0x7F) << 7));
    }

    private static int Color(XlsWriteContext context, ExcelColor color) =>
        color.IsAuto ? NoColor : context.StyleTable.ColorIndex(color);

    private static int Line(BorderLineStyle style) => style switch
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
}
