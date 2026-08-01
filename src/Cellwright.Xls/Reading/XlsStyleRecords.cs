using System.Buffers.Binary;

namespace Cellwright.Xls.Reading;

/// <summary>A FONT record, held until the XF table says which cells use it.</summary>
internal readonly record struct XlsFontRecord
{
    /// <summary>What a file with no usable FONT record gets.</summary>
    public static XlsFontRecord Default { get; } = new()
    {
        Name = "Arial",
        Size = 10,
        ColorIndex = 0x7FFF,
        Charset = -1,
    };

    /// <summary>Typeface name.</summary>
    public string Name { get; init; }

    /// <summary>Size in points.</summary>
    public double Size { get; init; }

    /// <summary>Whether the weight is bold or heavier.</summary>
    public bool Bold { get; init; }

    /// <summary>Whether the face is italic.</summary>
    public bool Italic { get; init; }

    /// <summary>Whether the text is struck through.</summary>
    public bool Strikethrough { get; init; }

    /// <summary>How the text is underlined.</summary>
    public UnderlineStyle Underline { get; init; }

    /// <summary>Whether the text is raised or lowered.</summary>
    public FontVerticalAlign VerticalAlign { get; init; }

    /// <summary>Palette slot of the text color.</summary>
    public int ColorIndex { get; init; }

    /// <summary>Font family class, for substitution when the face is missing.</summary>
    public int Family { get; init; }

    /// <summary>Windows character set number.</summary>
    public int Charset { get; init; }
}

/// <summary>
/// An XF record — one entry of the format table every cell points into. Parsed into a neutral
/// shape here, then turned into a <see cref="CellStyle"/> once the fonts and formats it refers
/// to have all been read.
/// </summary>
internal readonly record struct XlsXfRecord
{
    /// <summary>Index into the FONT records.</summary>
    public int FontIndex { get; init; }

    /// <summary>Number format id, builtin or custom.</summary>
    public int FormatIndex { get; init; }

    /// <summary>Whether the cell resists editing once the sheet is protected.</summary>
    public bool Locked { get; init; }

    /// <summary>Whether the formula bar stays empty for the cell on a protected sheet.</summary>
    public bool FormulaHidden { get; init; }

    /// <summary>Horizontal alignment.</summary>
    public HorizontalAlignment Horizontal { get; init; }

    /// <summary>Vertical alignment.</summary>
    public VerticalAlignment Vertical { get; init; }

    /// <summary>Whether long text wraps.</summary>
    public bool WrapText { get; init; }

    /// <summary>Whether the font shrinks to fit the column.</summary>
    public bool ShrinkToFit { get; init; }

    /// <summary>Indent steps.</summary>
    public int Indent { get; init; }

    /// <summary>Rotation in the model's convention.</summary>
    public int TextRotation { get; init; }

    /// <summary>Which way the text runs.</summary>
    public ReadingOrder ReadingOrder { get; init; }

    /// <summary>Fill pattern.</summary>
    public FillPattern FillPattern { get; init; }

    /// <summary>Palette slot of the pattern's foreground.</summary>
    public int ForegroundColorIndex { get; init; }

    /// <summary>Palette slot of the pattern's background.</summary>
    public int BackgroundColorIndex { get; init; }

    /// <summary>Left edge line style.</summary>
    public BorderLineStyle Left { get; init; }

    /// <summary>Right edge line style.</summary>
    public BorderLineStyle Right { get; init; }

    /// <summary>Top edge line style.</summary>
    public BorderLineStyle Top { get; init; }

    /// <summary>Bottom edge line style.</summary>
    public BorderLineStyle Bottom { get; init; }

    /// <summary>Diagonal line style.</summary>
    public BorderLineStyle Diagonal { get; init; }

    /// <summary>Palette slot of the left edge.</summary>
    public int LeftColorIndex { get; init; }

    /// <summary>Palette slot of the right edge.</summary>
    public int RightColorIndex { get; init; }

    /// <summary>Palette slot of the top edge.</summary>
    public int TopColorIndex { get; init; }

    /// <summary>Palette slot of the bottom edge.</summary>
    public int BottomColorIndex { get; init; }

    /// <summary>Palette slot of the diagonal.</summary>
    public int DiagonalColorIndex { get; init; }

    /// <summary>Whether the diagonal runs bottom-left to top-right.</summary>
    public bool DiagonalUp { get; init; }

    /// <summary>Whether the diagonal runs top-left to bottom-right.</summary>
    public bool DiagonalDown { get; init; }

    /// <summary>Parses the BIFF8 shape (MS-XLS 2.4.353), which is 20 bytes of dense bit fields.</summary>
    public static XlsXfRecord ParseBiff8(ReadOnlySpan<byte> body)
    {
        int fontIndex = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int formatIndex = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        ushort protection = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        byte alignment = body[6];
        byte rotation = body[7];
        byte indentation = body[8];
        ushort borders1 = BinaryPrimitives.ReadUInt16LittleEndian(body[10..]);
        ushort borders2 = BinaryPrimitives.ReadUInt16LittleEndian(body[12..]);
        uint borders3 = BinaryPrimitives.ReadUInt32LittleEndian(body[14..]);
        ushort fill = BinaryPrimitives.ReadUInt16LittleEndian(body[18..]);

        var diagonalLine = (BorderLineStyle)0;
        int diagonalFlags = (indentation >> 4) & 0x03;
        if (diagonalFlags != 0)
            diagonalLine = MapLine((byte)((borders3 >> 21) & 0x0F));

        return new XlsXfRecord
        {
            FontIndex = fontIndex,
            FormatIndex = formatIndex,
            Locked = (protection & 0x0001) != 0,
            FormulaHidden = (protection & 0x0002) != 0,
            Horizontal = MapHorizontal(alignment & 0x07),
            WrapText = (alignment & 0x08) != 0,
            Vertical = MapVertical((alignment >> 4) & 0x07),
            TextRotation = MapRotation(rotation),
            Indent = indentation & 0x0F,
            ShrinkToFit = (indentation & 0x10) != 0,
            ReadingOrder = (ReadingOrder)((indentation >> 6) & 0x03),
            Left = MapLine((byte)(borders1 & 0x0F)),
            Right = MapLine((byte)((borders1 >> 4) & 0x0F)),
            Top = MapLine((byte)((borders1 >> 8) & 0x0F)),
            Bottom = MapLine((byte)((borders1 >> 12) & 0x0F)),
            LeftColorIndex = borders2 & 0x7F,
            RightColorIndex = (borders2 >> 7) & 0x7F,
            DiagonalUp = (borders2 & 0x4000) != 0,
            DiagonalDown = (borders2 & 0x8000) != 0,
            TopColorIndex = (int)(borders3 & 0x7F),
            BottomColorIndex = (int)((borders3 >> 7) & 0x7F),
            DiagonalColorIndex = (int)((borders3 >> 14) & 0x7F),
            Diagonal = diagonalLine,
            FillPattern = MapPattern((byte)((borders3 >> 26) & 0x3F)),
            ForegroundColorIndex = fill & 0x7F,
            BackgroundColorIndex = (fill >> 7) & 0x7F,
        };
    }

    /// <summary>
    /// Parses the older shapes. They agree on the first four bytes and diverge after, so only
    /// the fields that carry the same meaning everywhere are taken; the rest default.
    /// </summary>
    public static XlsXfRecord ParseLegacy(ReadOnlySpan<byte> body, BiffVersion version)
    {
        if (body.Length < 4)
            return new XlsXfRecord { FormatIndex = 0 };

        if (version <= BiffVersion.Biff2)
        {
            // BIFF2: font index, then a format-and-style byte, then two of cell protection.
            return new XlsXfRecord
            {
                FontIndex = body[0],
                FormatIndex = body[2] & 0x3F,
                Locked = (body[2] & 0x40) != 0,
                FormulaHidden = (body[2] & 0x80) != 0,
            };
        }

        int fontIndex = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int formatIndex = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        if (body.Length < 16)
            return new XlsXfRecord { FontIndex = fontIndex, FormatIndex = formatIndex };

        ushort protection = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        byte alignment = body[6];
        ushort fill = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);
        ushort borders1 = BinaryPrimitives.ReadUInt16LittleEndian(body[10..]);
        ushort borders2 = BinaryPrimitives.ReadUInt16LittleEndian(body[12..]);

        return new XlsXfRecord
        {
            FontIndex = fontIndex,
            FormatIndex = formatIndex,
            Locked = (protection & 0x0001) != 0,
            FormulaHidden = (protection & 0x0002) != 0,
            Horizontal = MapHorizontal(alignment & 0x07),
            WrapText = (alignment & 0x08) != 0,
            Vertical = MapVertical((alignment >> 4) & 0x07),
            FillPattern = MapPattern((byte)(fill & 0x3F)),
            ForegroundColorIndex = (fill >> 6) & 0x7F,
            BackgroundColorIndex = (fill >> 13) & 0x7F,
            Bottom = MapLine((byte)(borders1 & 0x07)),
            BottomColorIndex = (borders1 >> 3) & 0x7F,
            Top = MapLine((byte)((borders1 >> 10) & 0x07)),
            Left = MapLine((byte)((borders1 >> 13) & 0x07)),
            TopColorIndex = borders2 & 0x7F,
            Right = MapLine((byte)((borders2 >> 7) & 0x07)),
            RightColorIndex = (borders2 >> 10) & 0x7F,
        };
    }

    private static HorizontalAlignment MapHorizontal(int value) => value switch
    {
        1 => HorizontalAlignment.Left,
        2 => HorizontalAlignment.Center,
        3 => HorizontalAlignment.Right,
        4 => HorizontalAlignment.Fill,
        5 => HorizontalAlignment.Justify,
        6 => HorizontalAlignment.CenterContinuous,
        7 => HorizontalAlignment.Distributed,
        _ => HorizontalAlignment.General,
    };

    private static VerticalAlignment MapVertical(int value) => value switch
    {
        0 => VerticalAlignment.Top,
        1 => VerticalAlignment.Center,
        3 => VerticalAlignment.Justify,
        4 => VerticalAlignment.Distributed,
        _ => VerticalAlignment.Bottom,
    };

    /// <summary>
    /// Maps BIFF's rotation byte onto the model's. They agree from 0 to 90; BIFF says 91-180 for
    /// clockwise, which the model repeats, and 255 for stacked, which it also repeats.
    /// </summary>
    private static int MapRotation(byte value) => value switch
    {
        255 => Alignment.StackedRotation,
        <= 180 => value,
        _ => 0,
    };

    private static BorderLineStyle MapLine(byte value) => value switch
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

    private static FillPattern MapPattern(byte value) => value switch
    {
        1 => FillPattern.Solid,
        2 => FillPattern.MediumGray,
        3 => FillPattern.DarkGray,
        4 => FillPattern.LightGray,
        5 => FillPattern.DarkHorizontal,
        6 => FillPattern.DarkVertical,
        7 => FillPattern.DarkDown,
        8 => FillPattern.DarkUp,
        9 => FillPattern.DarkGrid,
        10 => FillPattern.DarkTrellis,
        11 => FillPattern.LightHorizontal,
        12 => FillPattern.LightVertical,
        13 => FillPattern.LightDown,
        14 => FillPattern.LightUp,
        15 => FillPattern.LightGrid,
        16 => FillPattern.LightTrellis,
        17 => FillPattern.Gray125,
        18 => FillPattern.Gray0625,
        _ => FillPattern.None,
    };
}
