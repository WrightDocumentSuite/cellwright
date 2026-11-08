using System.Buffers.Binary;
using Cellwright.Styles;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the styles part into the workbook's style pool.
/// </summary>
/// <remarks>
/// The tables here are the same four the older format has — number formats, fonts, fills,
/// borders — but colours are written outright as ARGB instead of as slots in a palette of 56,
/// which is the one thing that makes this simpler rather than harder than its predecessor.
/// Differential formats are a fifth, property-sized table returned for conditional rules to
/// attach after their sheet has been parsed.
/// </remarks>
internal static class XlsbStyles
{
    /// <returns>Differential formats in <c>dxfId</c> order.</returns>
    public static IReadOnlyList<DifferentialStyle> Read(byte[] part, XlsGlobals globals, StylePool pool)
    {
        var formats = new Dictionary<int, string>();
        var fonts = new List<FontStyle>();
        var fills = new List<FillStyle>();
        var borders = new List<BorderStyle>();
        var cellFormats = new List<XlsbXf>();
        var differentialStyles = new List<DifferentialStyle>();

        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            switch (reader.RecordType)
            {
                case Biff12RecordType.Fmt:
                    ReadFormat(reader.Record, formats);
                    break;
                case Biff12RecordType.Font:
                    fonts.Add(ReadFont(reader.Record));
                    break;
                case Biff12RecordType.Fill:
                    fills.Add(ReadFill(reader.Record));
                    break;
                case Biff12RecordType.Border:
                    borders.Add(ReadBorder(reader.Record));
                    break;

                // Both tables are made of BrtXF records and only the surrounding markers tell
                // them apart; a cell's index counts into the second one.
                case Biff12RecordType.BeginCellStyleXFs:
                    reader.SkipCollection(Biff12RecordType.BeginCellStyleXFs, Biff12RecordType.EndCellStyleXFs);
                    break;
                case Biff12RecordType.Xf:
                    cellFormats.Add(ReadXf(reader.Record));
                    break;
                case Biff12RecordType.Dxf:
                    // A malformed entry still occupies its position in the table: dxfId is an
                    // index, so dropping one would silently attach every later style to the
                    // wrong conditional-formatting rule.
                    differentialStyles.Add(ReadDifferentialStyle(reader.Record, formats));
                    break;

                default:
                    break;
            }
        }

        globals.Fonts.AddRange(fonts);

        foreach (XlsbXf xf in cellFormats)
        {
            string code = Resolve(formats, xf.Format);
            globals.CellFormats.Add(new StyleId(pool.Intern(Build(xf, code, fonts, fills, borders))));
            globals.TemporalKinds.Add(
                NumberFormatClassifier.Classify(xf.Format, code.Length == 0 ? null : code));
        }

        return differentialStyles;
    }

    /// <summary>One entry of the cell format table, before the tables it points into are joined.</summary>
    private readonly record struct XlsbXf(int Format, int Font, int Fill, int Border, bool Locked, bool Hidden);

    private static string Resolve(IReadOnlyDictionary<int, string> formats, int index) =>
        formats.TryGetValue(index, out string? custom) ? custom : BuiltinNumberFormats.GetCode(index) ?? "";

    private static void ReadFormat(ReadOnlySpan<byte> body, Dictionary<int, string> formats)
    {
        if (body.Length < 2)
            return;

        int index = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int offset = 2;
        formats[index] = Biff12Values.ReadString(body, ref offset);
    }

    private static FontStyle ReadFont(ReadOnlySpan<byte> body)
    {
        if (body.Length < 21)
            return FontStyle.Default;

        int twips = BinaryPrimitives.ReadUInt16LittleEndian(body);
        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        int weight = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int script = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        byte underline = body[8];
        ExcelColor color = ReadColor(body[12..]);

        int offset = 21;                                    // Past brtColor and bFontScheme.
        string name = Biff12Values.ReadString(body, ref offset);

        return new FontStyle
        {
            Name = name.Length == 0 ? FontStyle.Default.Name : name,
            Size = twips / 20.0,
            Bold = weight >= 0x02BC,
            Italic = (flags & 0x02) != 0,
            Strikethrough = (flags & 0x08) != 0,
            Underline = underline switch
            {
                0x01 => UnderlineStyle.Single,
                0x02 => UnderlineStyle.Double,
                0x21 => UnderlineStyle.SingleAccounting,
                0x22 => UnderlineStyle.DoubleAccounting,
                _ => UnderlineStyle.None,
            },
            VerticalAlign = script switch
            {
                1 => FontVerticalAlign.Superscript,
                2 => FontVerticalAlign.Subscript,
                _ => FontVerticalAlign.Baseline,
            },
            Color = color,
        };
    }

    private static FillStyle ReadFill(ReadOnlySpan<byte> body)
    {
        if (body.Length < 20)
            return FillStyle.None;

        uint pattern = BinaryPrimitives.ReadUInt32LittleEndian(body);
        FillPattern mapped = MapPattern(pattern);
        return mapped == FillPattern.None
            ? FillStyle.None
            : FillStyle.Patterned(mapped, ReadColor(body[4..]), ReadColor(body[12..]));
    }

    private static BorderStyle ReadBorder(ReadOnlySpan<byte> body)
    {
        // Flags byte, then five edges of two bytes of style plus eight of colour.
        if (body.Length < 2 + (5 * 10))
            return BorderStyle.None;

        byte flags = body[0];
        return new BorderStyle
        {
            Top = ReadEdge(body, 0),
            Bottom = ReadEdge(body, 1),
            Left = ReadEdge(body, 2),
            Right = ReadEdge(body, 3),
            Diagonal = ReadEdge(body, 4),
            DiagonalUp = (flags & 0x01) != 0,
            DiagonalDown = (flags & 0x02) != 0,
        };
    }

    private static BorderEdge ReadEdge(ReadOnlySpan<byte> body, int index)
    {
        int offset = 2 + (index * 10);
        return new BorderEdge(MapLine(body[offset]), ReadColor(body[(offset + 2)..]));
    }

    private static XlsbXf ReadXf(ReadOnlySpan<byte> body)
    {
        if (body.Length < 16)
            return default;

        return new XlsbXf(
            BinaryPrimitives.ReadUInt16LittleEndian(body[2..]),
            BinaryPrimitives.ReadUInt16LittleEndian(body[4..]),
            BinaryPrimitives.ReadUInt16LittleEndian(body[6..]),
            BinaryPrimitives.ReadUInt16LittleEndian(body[8..]),
            (BinaryPrimitives.ReadUInt16LittleEndian(body[14..]) & 0x8000) == 0,
            false);
    }

    private static CellStyle Build(
        in XlsbXf xf, string numberFormat,
        List<FontStyle> fonts, List<FillStyle> fills, List<BorderStyle> borders) => new()
        {
            NumberFormat = numberFormat,
            Font = (uint)xf.Font < (uint)fonts.Count ? fonts[xf.Font] : FontStyle.Default,
            Fill = (uint)xf.Fill < (uint)fills.Count ? fills[xf.Fill] : FillStyle.None,
            Border = (uint)xf.Border < (uint)borders.Count ? borders[xf.Border] : BorderStyle.None,
        };

    /// <summary>
    /// Reads a <c>BrtDXF</c> (MS-XLSB 2.4.359). Its <c>XFProps</c> array is deliberately
    /// property-sized: unsupported additions can be skipped by <c>cb</c> without losing the
    /// boundary of the next property.
    /// </summary>
    private static DifferentialStyle ReadDifferentialStyle(
        ReadOnlySpan<byte> body, IReadOnlyDictionary<int, string> formats)
    {
        // Two bytes of BrtDXF flags, then XFProps.reserved and XFProps.cprops.
        if (body.Length < 6)
            return DifferentialStyle.Empty;

        var style = DifferentialStyle.Empty;
        ExcelColor foreground = ExcelColor.Auto;
        ExcelColor background = ExcelColor.Auto;
        BorderStyle? border = null;
        string? directNumberFormat = null;
        int? numberFormatId = null;
        int propertyCount = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int offset = 6;

        for (int index = 0; index < propertyCount; index++)
        {
            if (offset + 4 > body.Length)
                break;

            int type = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
            int size = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]);
            if (size < 4 || size > body.Length - offset)
                break;

            ReadOnlySpan<byte> data = body.Slice(offset + 4, size - 4);
            offset += size;

            switch (type)
            {
                case 0x0001: // foreground fill color
                    foreground = ReadPropertyColor(data);
                    break;
                case 0x0002: // background fill color
                    background = ReadPropertyColor(data);
                    break;
                case 0x0005: // text color
                    style = style with { FontColor = ReadPropertyColor(data) };
                    break;
                case 0x0006: // top border
                    border = (border ?? BorderStyle.None) with { Top = ReadPropertyBorder(data) };
                    break;
                case 0x0007: // bottom border
                    border = (border ?? BorderStyle.None) with { Bottom = ReadPropertyBorder(data) };
                    break;
                case 0x0008: // left border
                    border = (border ?? BorderStyle.None) with { Left = ReadPropertyBorder(data) };
                    break;
                case 0x0009: // right border
                    border = (border ?? BorderStyle.None) with { Right = ReadPropertyBorder(data) };
                    break;
                case 0x000A: // diagonal border
                    border = (border ?? BorderStyle.None) with { Diagonal = ReadPropertyBorder(data) };
                    break;
                case 0x000D when data.Length >= 1:
                    border = (border ?? BorderStyle.None) with { DiagonalUp = data[0] != 0 };
                    break;
                case 0x000E when data.Length >= 1:
                    border = (border ?? BorderStyle.None) with { DiagonalDown = data[0] != 0 };
                    break;
                case 0x0019 when data.Length >= 2: // Bold
                    style = style with { Bold = BinaryPrimitives.ReadUInt16LittleEndian(data) >= 0x02BC };
                    break;
                case 0x001A when data.Length >= 2: // Underline; the model has a single toggle.
                    style = style with { Underline = BinaryPrimitives.ReadUInt16LittleEndian(data) != 0 };
                    break;
                case 0x001C when data.Length >= 1:
                    style = style with { Italic = data[0] != 0 };
                    break;
                case 0x001D when data.Length >= 1:
                    style = style with { Strikethrough = data[0] != 0 };
                    break;
                case 0x0026 when TryReadNumberFormat(data, out string? code):
                    directNumberFormat = code;
                    break;
                case 0x0029 when data.Length >= 2:
                    numberFormatId = BinaryPrimitives.ReadUInt16LittleEndian(data);
                    break;
                default:
                    // Alignment, gradients, font name/size and future properties are not in
                    // DifferentialStyle. The size field has already advanced over their blob.
                    break;
            }
        }

        // Like the OOXML dxf reader, prefer the background spelling. Excel's own common
        // conditional-format preset writes only property 0x0002 and no FillPattern property.
        ExcelColor paint = !background.IsAuto ? background : foreground;
        if (!paint.IsAuto)
            style = style with { FillColor = paint };
        if (border is not null)
            style = style with { Border = border };
        if (directNumberFormat is not null)
            style = style with { NumberFormat = directNumberFormat };
        else if (numberFormatId is int formatId)
            style = style with { NumberFormat = Resolve(formats, formatId) };
        return style;
    }

    /// <summary>An XFPropColor packs its validity bit before the seven-bit color kind.</summary>
    private static ExcelColor ReadPropertyColor(ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return ExcelColor.Auto;

        int type = body[0] >> 1;
        int index = body[1];
        short rawTint = BinaryPrimitives.ReadInt16LittleEndian(body[2..]);
        double tint = rawTint == short.MinValue ? 0 : rawTint / 32767.0;
        return type switch
        {
            1 => ExcelColor.FromIndexed(index),
            2 => ExcelColor.FromArgb((uint)((body[7] << 24) | (body[4] << 16) | (body[5] << 8) | body[6])),
            3 => ExcelColor.FromTheme(SwapDarkAndLight(index), tint),
            _ => ExcelColor.Auto,
        };
    }

    private static BorderEdge ReadPropertyBorder(ReadOnlySpan<byte> body) => body.Length < 10
        ? default
        : new BorderEdge(MapLine(BinaryPrimitives.ReadUInt16LittleEndian(body[8..])), ReadPropertyColor(body));

    /// <summary>The direct format-string property is an LPWideString.</summary>
    private static bool TryReadNumberFormat(ReadOnlySpan<byte> body, out string? code)
    {
        code = null;
        if (body.Length < 2)
            return false;

        int characters = BinaryPrimitives.ReadUInt16LittleEndian(body);
        if (characters <= 0 || characters * 2 != body.Length - 2)
            return false;

        code = System.Text.Encoding.Unicode.GetString(body[2..]);
        return true;
    }

    /// <summary>
    /// Reads a <c>BrtColor</c> (MS-XLSB 2.4.337): a type, an index, a tint, and four channels.
    /// Which of them mean anything depends on the type, and only one of the four kinds carries
    /// its colour in the channels. Shared with the conditional formatting reader, whose graded
    /// rules state their colors in the same record.
    /// </summary>
    internal static ExcelColor ReadColor(ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return ExcelColor.Auto;

        int packedType = body[0];
        bool validRgb = (packedType & 0x01) != 0;
        int type = packedType >> 1;
        int index = body[1];

        // The tint is a fraction of the full range, which the format spends a signed short on:
        // its extremes are complete lightening and complete darkening.
        short rawTint = BinaryPrimitives.ReadInt16LittleEndian(body[2..]);
        double tint = rawTint < 0 ? rawTint / 32768.0 : rawTint / 32767.0;

        return type switch
        {
            1 => ExcelColor.FromIndexed(index),
            // fValidRGB is the low bit preceding xColorType, and MUST be one for an
            // explicit RGB color. Treat a producer that violates that invariant as not
            // having supplied a usable color instead of consuming undefined channels.
            2 when validRgb => ExcelColor.FromArgb(
                (uint)((body[7] << 24) | (body[4] << 16) | (body[5] << 8) | body[6])),

            // The theme numbers its slots as the scheme states them, dark before light; a style
            // names the same slots with the first two pairs swapped, which is what the model
            // holds, so the pair has to be exchanged on the way in.
            3 => ExcelColor.FromTheme(SwapDarkAndLight(index), tint),
            _ => ExcelColor.Auto,
        };
    }

    /// <summary>Exchanges dk1 with lt1 and dk2 with lt2, leaving the rest of the slots alone.</summary>
    private static int SwapDarkAndLight(int index) => index switch
    {
        0 => 1,
        1 => 0,
        2 => 3,
        3 => 2,
        _ => index,
    };

    private static FillPattern MapPattern(uint pattern) => pattern switch
    {
        0 => FillPattern.None,
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

    private static BorderLineStyle MapLine(ushort style) => style switch
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
