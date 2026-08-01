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
/// </remarks>
internal static class XlsbStyles
{
    public static void Read(byte[] part, XlsGlobals globals, StylePool pool)
    {
        var formats = new Dictionary<int, string>();
        var fonts = new List<FontStyle>();
        var fills = new List<FillStyle>();
        var borders = new List<BorderStyle>();
        var cellFormats = new List<XlsbXf>();

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

                default:
                    break;
            }
        }

        foreach (XlsbXf xf in cellFormats)
        {
            string code = Resolve(formats, xf.Format);
            globals.CellFormats.Add(new StyleId(pool.Intern(Build(xf, code, fonts, fills, borders))));
            globals.TemporalKinds.Add(
                NumberFormatClassifier.Classify(xf.Format, code.Length == 0 ? null : code));
        }
    }

    /// <summary>One entry of the cell format table, before the tables it points into are joined.</summary>
    private readonly record struct XlsbXf(int Format, int Font, int Fill, int Border, bool Locked, bool Hidden);

    private static string Resolve(Dictionary<int, string> formats, int index) =>
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
    /// Reads a <c>BrtColor</c> (MS-XLSB 2.4.337): a type, an index, a tint, and four channels.
    /// Which of them mean anything depends on the type, and only one of the four kinds carries
    /// its colour in the channels. Shared with the conditional formatting reader, whose graded
    /// rules state their colors in the same record.
    /// </summary>
    internal static ExcelColor ReadColor(ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return ExcelColor.Auto;

        int type = body[0] & 0x7F;
        int index = body[1];

        // The tint is a fraction of the full range, which the format spends a signed short on:
        // its extremes are complete lightening and complete darkening.
        double tint = BinaryPrimitives.ReadInt16LittleEndian(body[2..]) / 32767.0;

        return type switch
        {
            1 => ExcelColor.FromIndexed(index),
            2 => ExcelColor.FromArgb((uint)((body[7] << 24) | (body[4] << 16) | (body[5] << 8) | body[6])),

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

    private static BorderLineStyle MapLine(byte style) => style switch
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
