using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Styles;

/// <summary>
/// Parsed styles part: full <see cref="CellStyle"/> records per cell-format index plus the
/// precomputed temporal classification used to turn serial numbers into dates and durations.
/// </summary>
internal sealed record StylesPart(IReadOnlyList<CellStyle> CellFormats, TemporalKind[] TemporalKinds)
{
    /// <summary>An empty styles part (workbook without styles.xml).</summary>
    public static StylesPart Empty { get; } = new([], []);

    /// <summary>Sections of the part the model does not represent, carried through on save.</summary>
    public StylesMarkup Markup { get; init; } = new();

    /// <summary>
    /// The differential formats of <c>dxfs</c>, parsed into the model's partial-style records by
    /// position, so a conditional rule read from the file can show what it paints with. Writing
    /// still replays the raw markup — this list is the readable view, not the storage.
    /// </summary>
    public IReadOnlyList<DifferentialStyle> DifferentialStyles { get; init; } = [];

    /// <summary>Parses a buffered styles.xml stream.</summary>
    public static StylesPart Parse(Stream stream)
    {
        var parser = new StylesPartParser();
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        return parser.Parse(xml);
    }
}

/// <summary>
/// Single-pass parser for styles.xml. Component tables (fonts, fills, borders, custom
/// number formats) are read first, then cell formats are composed into records.
/// </summary>
internal sealed class StylesPartParser
{
    private readonly Dictionary<int, string> _customFormats = [];
    private readonly List<FontStyle> _fonts = [];
    private readonly List<FillStyle> _fills = [];
    private readonly List<BorderStyle> _borders = [];
    private readonly List<CellStyle> _cellFormats = [];
    private readonly List<CellStyle> _styleFormats = [];
    private readonly List<int> _formatIds = [];
    private readonly StylesMarkup _markup = new();

    public StylesPart Parse(XmlReader xml)
    {
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "styleSheet":
                    _markup.RootAttributes = RawXml.ReadAttributes(xml, "xmlns");
                    break;

                // Scoped deliberately: differential formats carry numFmt elements of their own,
                // and letting those into the workbook table silently retypes dated columns.
                case "numFmts":
                    ParseSection(xml, "numFmts", "numFmt", static (p, r) => p.ParseNumberFormat(r));
                    break;
                case "fonts":
                    ParseSection(xml, "fonts", "font", static (p, r) => p._fonts.Add(p.ParseFont(r)));
                    break;
                case "fills":
                    ParseSection(xml, "fills", "fill", static (p, r) => p._fills.Add(p.ParseFill(r)));
                    break;
                case "borders":
                    ParseSection(xml, "borders", "border", static (p, r) => p._borders.Add(p.ParseBorder(r)));
                    break;
                case "cellStyleXfs":
                    ParseSection(xml, "cellStyleXfs", "xf", static (p, r) => p.ParseCellFormat(r, p._styleFormats));
                    break;
                case "cellXfs":
                    ParseSection(xml, "cellXfs", "xf", static (p, r) => p.ParseCellFormat(r, p._cellFormats));
                    break;
                case "cellStyles":
                    _markup.CellStyleEntries = xml.ReadOuterXml();
                    advance = false;
                    break;
                // Kept as inner markup rather than as a whole element: conditional formatting
                // added by the model appends its own formats after these, and the count on the
                // wrapper has to grow with them.
                case "dxfs":
                    _markup.DifferentialFormatCount =
                        int.TryParse(xml.GetAttribute("count"), NumberStyles.None, CultureInfo.InvariantCulture, out int dxfCount)
                            ? dxfCount
                            : 0;
                    _markup.DifferentialFormats = xml.ReadInnerXml();
                    advance = false;
                    break;
                case "tableStyles":
                    _markup.TableStyles = xml.ReadOuterXml();
                    advance = false;
                    break;
                case "colors":
                    _markup.Colors = xml.ReadOuterXml();
                    advance = false;
                    break;
                case "extLst":
                    _markup.Extensions = xml.ReadOuterXml();
                    advance = false;
                    break;
                default:
                    break;
            }
        }

        var temporal = new TemporalKind[_cellFormats.Count];
        for (int i = 0; i < temporal.Length; i++)
            temporal[i] = NumberFormatClassifier.Classify(_formatIds[i], _customFormats.GetValueOrDefault(_formatIds[i]));

        if (_styleFormats.Count > 0)
            _markup.StyleFormats = _styleFormats;

        return new StylesPart(_cellFormats, temporal)
        {
            Markup = _markup,
            DifferentialStyles = ParseDifferentialFormats(_markup.DifferentialFormats),
        };
    }

    private void ParseNumberFormat(XmlReader xml)
    {
        if (int.TryParse(xml.GetAttribute("numFmtId"), NumberStyles.None, CultureInfo.InvariantCulture, out int id) &&
            xml.GetAttribute("formatCode") is { } code)
            _customFormats[id] = code;
    }

    /// <summary>Iterates the direct item elements of a container section.</summary>
    private void ParseSection(XmlReader xml, string container, string item, Action<StylesPartParser, XmlReader> onItem)
    {
        if (xml.IsEmptyElement)
            return;

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == container)
                break;
            if (xml.NodeType == XmlNodeType.Element && xml.LocalName == item)
                onItem(this, xml);
        }
    }

    /// <summary>Parses a <c>font</c> element; the reader ends on its end tag (or the empty element itself).</summary>
    private FontStyle ParseFont(XmlReader xml) => FontMarkup.Parse(xml, "font");

    /// <summary>Parses a <c>fill</c>: either a <c>patternFill</c> or a <c>gradientFill</c>.</summary>
    private FillStyle ParseFill(XmlReader xml)
    {
        if (xml.IsEmptyElement)
            return FillStyle.None;

        FillPattern pattern = FillPattern.None;
        ExcelColor foreground = ExcelColor.Auto;
        ExcelColor background = ExcelColor.Auto;
        GradientFill? gradient = null;
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "fill")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "patternFill":
                    pattern = ParsePattern(xml.GetAttribute("patternType"));
                    break;
                case "fgColor":
                    foreground = ParseColor(xml);
                    break;
                case "bgColor":
                    background = ParseColor(xml);
                    break;
                case "gradientFill":
                    gradient = ParseGradient(xml);
                    break;
                default:
                    break;
            }
        }

        if (gradient is not null)
            return new FillStyle { Gradient = gradient };

        return pattern == FillPattern.None
            ? FillStyle.None
            : new FillStyle { Pattern = pattern, Color = foreground, BackgroundColor = background };
    }

    /// <summary>Parses a <c>gradientFill</c> (§18.8.24): the sweep attributes and the color stops.</summary>
    private static GradientFill ParseGradient(XmlReader xml)
    {
        var gradient = new GradientFill
        {
            Type = xml.GetAttribute("type") == "path" ? GradientFillType.Path : GradientFillType.Linear,
            Degree = ParseDoubleAttribute(xml, "degree"),
            Left = ParseDoubleAttribute(xml, "left"),
            Right = ParseDoubleAttribute(xml, "right"),
            Top = ParseDoubleAttribute(xml, "top"),
            Bottom = ParseDoubleAttribute(xml, "bottom"),
        };

        if (xml.IsEmptyElement)
            return gradient;

        List<GradientStop> stops = [];
        double position = 0;
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "gradientFill")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "stop")
                position = ParseDoubleAttribute(xml, "position");
            else if (xml.LocalName == "color")
                stops.Add(new GradientStop(position, ParseColor(xml)));
        }

        return gradient with { Stops = stops };
    }

    private BorderStyle ParseBorder(XmlReader xml)
    {
        BorderStyle border = BorderStyle.None;
        if (xml.IsEmptyElement)
            return border;

        // The two flags sit on the border element itself while the line they describe is a
        // child, so they are read before the loop reaches it.
        border = border with
        {
            DiagonalUp = xml.GetAttribute("diagonalUp") is "1" or "true",
            DiagonalDown = xml.GetAttribute("diagonalDown") is "1" or "true",
        };

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "border")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            string edgeName = xml.LocalName;
            if (edgeName is not ("left" or "right" or "top" or "bottom" or "diagonal"))
                continue;

            BorderEdge edge = ParseEdge(xml, edgeName);
            border = edgeName switch
            {
                "left" => border with { Left = edge },
                "right" => border with { Right = edge },
                "top" => border with { Top = edge },
                "diagonal" => border with { Diagonal = edge },
                _ => border with { Bottom = edge },
            };
        }

        return border;
    }

    private BorderEdge ParseEdge(XmlReader xml, string edgeName)
    {
        BorderLineStyle line = xml.GetAttribute("style") switch
        {
            "thin" => BorderLineStyle.Thin,
            "medium" => BorderLineStyle.Medium,
            "thick" => BorderLineStyle.Thick,
            "dashed" => BorderLineStyle.Dashed,
            "dotted" => BorderLineStyle.Dotted,
            "double" => BorderLineStyle.Double,
            "hair" => BorderLineStyle.Hair,
            "mediumDashed" => BorderLineStyle.MediumDashed,
            "dashDot" => BorderLineStyle.DashDot,
            "mediumDashDot" => BorderLineStyle.MediumDashDot,
            "dashDotDot" => BorderLineStyle.DashDotDot,
            "mediumDashDotDot" => BorderLineStyle.MediumDashDotDot,
            "slantDashDot" => BorderLineStyle.SlantDashDot,
            null or "none" => BorderLineStyle.None,
            _ => BorderLineStyle.Thin,
        };

        ExcelColor color = ExcelColor.Auto;
        if (!xml.IsEmptyElement)
        {
            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == edgeName)
                    break;
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "color")
                    color = ParseColor(xml);
            }
        }

        return new BorderEdge(line, color);
    }

    /// <summary>
    /// Parses one <c>xf</c> into the given table — the cell formats of <c>cellXfs</c> or the
    /// masters of <c>cellStyleXfs</c>. Only cell formats feed the temporal classification;
    /// nothing ever displays a value through a master directly.
    /// </summary>
    private void ParseCellFormat(XmlReader xml, List<CellStyle> target)
    {
        int numberFormatId = ParseIntAttribute(xml, "numFmtId");
        int fontId = ParseIntAttribute(xml, "fontId");
        int fillId = ParseIntAttribute(xml, "fillId");
        int borderId = ParseIntAttribute(xml, "borderId");
        int styleXfId = ParseIntAttribute(xml, "xfId");

        Alignment alignment = Alignment.Default;
        CellProtection protection = CellProtection.Default;
        if (!xml.IsEmptyElement)
        {
            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "xf")
                    break;
                if (xml.NodeType != XmlNodeType.Element)
                    continue;

                if (xml.LocalName == "alignment")
                    alignment = ParseAlignment(xml);
                else if (xml.LocalName == "protection")
                    protection = ParseProtection(xml);
            }
        }

        string numberFormat = _customFormats.GetValueOrDefault(numberFormatId)
            ?? BuiltinNumberFormats.GetCode(numberFormatId)
            ?? "";

        if (ReferenceEquals(target, _cellFormats))
            _formatIds.Add(numberFormatId);
        target.Add(new CellStyle
        {
            Font = (uint)fontId < (uint)_fonts.Count ? _fonts[fontId] : FontStyle.Default,
            Fill = (uint)fillId < (uint)_fills.Count ? _fills[fillId] : FillStyle.None,
            Border = (uint)borderId < (uint)_borders.Count ? _borders[borderId] : BorderStyle.None,
            Alignment = alignment,
            Protection = protection,
            NumberFormat = numberFormat,
            StyleXfId = ReferenceEquals(target, _cellFormats) ? styleXfId : 0,
        });
    }

    /// <summary>Reads the <c>protection</c> child of an <c>xf</c>; both flags default the Excel way.</summary>
    private static CellProtection ParseProtection(XmlReader xml)
    {
        bool locked = xml.GetAttribute("locked") is not ("0" or "false");
        bool hidden = xml.GetAttribute("hidden") is "1" or "true";
        return locked && !hidden ? CellProtection.Default : new CellProtection { Locked = locked, Hidden = hidden };
    }

    private static Alignment ParseAlignment(XmlReader xml) => new()
    {
        Horizontal = xml.GetAttribute("horizontal") switch
        {
            "left" => HorizontalAlignment.Left,
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            "justify" => HorizontalAlignment.Justify,
            "fill" => HorizontalAlignment.Fill,
            "centerContinuous" => HorizontalAlignment.CenterContinuous,
            "distributed" => HorizontalAlignment.Distributed,
            _ => HorizontalAlignment.General,
        },
        Vertical = xml.GetAttribute("vertical") switch
        {
            "top" => VerticalAlignment.Top,
            "center" => VerticalAlignment.Center,
            "justify" => VerticalAlignment.Justify,
            "distributed" => VerticalAlignment.Distributed,
            _ => VerticalAlignment.Bottom,
        },
        WrapText = xml.GetAttribute("wrapText") is "1" or "true",
        Indent = ParseIntAttribute(xml, "indent"),
        TextRotation = ParseIntAttribute(xml, "textRotation"),
        ShrinkToFit = xml.GetAttribute("shrinkToFit") is "1" or "true",
        ReadingOrder = ParseIntAttribute(xml, "readingOrder") switch
        {
            1 => ReadingOrder.LeftToRight,
            2 => ReadingOrder.RightToLeft,
            _ => ReadingOrder.Context,
        },
    };

    internal static UnderlineStyle ParseUnderline(string? value) => value switch
    {
        null or "single" => UnderlineStyle.Single,
        "double" => UnderlineStyle.Double,
        "singleAccounting" => UnderlineStyle.SingleAccounting,
        "doubleAccounting" => UnderlineStyle.DoubleAccounting,
        _ => UnderlineStyle.None,
    };

    internal static FontVerticalAlign ParseVerticalAlign(string? value) => value switch
    {
        "superscript" => FontVerticalAlign.Superscript,
        "subscript" => FontVerticalAlign.Subscript,
        _ => FontVerticalAlign.Baseline,
    };

    internal static FontScheme ParseScheme(string? value) => value switch
    {
        "major" => FontScheme.Major,
        "minor" => FontScheme.Minor,
        _ => FontScheme.None,
    };

    private static FillPattern ParsePattern(string? value) => value switch
    {
        "solid" => FillPattern.Solid,
        "darkGray" => FillPattern.DarkGray,
        "mediumGray" => FillPattern.MediumGray,
        "lightGray" => FillPattern.LightGray,
        "gray125" => FillPattern.Gray125,
        "gray0625" => FillPattern.Gray0625,
        "darkHorizontal" => FillPattern.DarkHorizontal,
        "darkVertical" => FillPattern.DarkVertical,
        "darkDown" => FillPattern.DarkDown,
        "darkUp" => FillPattern.DarkUp,
        "darkGrid" => FillPattern.DarkGrid,
        "darkTrellis" => FillPattern.DarkTrellis,
        "lightHorizontal" => FillPattern.LightHorizontal,
        "lightVertical" => FillPattern.LightVertical,
        "lightDown" => FillPattern.LightDown,
        "lightUp" => FillPattern.LightUp,
        "lightGrid" => FillPattern.LightGrid,
        "lightTrellis" => FillPattern.LightTrellis,
        _ => FillPattern.None,
    };

    /// <summary>Boolean flag elements like <c>&lt;b/&gt;</c>: present means true unless val says otherwise.</summary>
    internal static bool IsFlagSet(XmlReader xml) => xml.GetAttribute("val") is not ("0" or "false");

    /// <summary>
    /// Reads a color element in any of its four spellings. A theme slot is kept as one, since
    /// the theme part travels with the package; the legacy indexed palette is resolved here,
    /// because its numbers mean nothing outside this table.
    /// </summary>
    internal static ExcelColor ParseColor(XmlReader xml)
    {
        if (xml.GetAttribute("auto") is "1" or "true")
            return ExcelColor.Auto;

        if (xml.GetAttribute("rgb") is { } rgb && ExcelColor.TryParseHex(rgb, out ExcelColor color))
            return color;

        if (int.TryParse(xml.GetAttribute("theme"), NumberStyles.None, CultureInfo.InvariantCulture, out int theme))
        {
            double tint = double.TryParse(xml.GetAttribute("tint"), CultureInfo.InvariantCulture, out double parsed) ? parsed : 0;
            return ExcelColor.FromTheme(theme, tint);
        }

        if (int.TryParse(xml.GetAttribute("indexed"), NumberStyles.None, CultureInfo.InvariantCulture, out int indexed))
            return ExcelColor.FromIndexed(indexed);

        return ExcelColor.Auto;
    }

    internal static int ParseIntAttribute(XmlReader xml, string name) =>
        int.TryParse(xml.GetAttribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out int value) ? value : 0;

    private static double ParseDoubleAttribute(XmlReader xml, string name) =>
        double.TryParse(xml.GetAttribute(name), CultureInfo.InvariantCulture, out double value) ? value : 0;

    /// <summary>
    /// Parses the captured <c>dxfs</c> content into partial styles, one per <c>dxf</c> in file
    /// order. The raw markup stays what gets written back; this is the view a conditional rule
    /// shows through <c>Style</c>. A format the model cannot express parses to what it can.
    /// </summary>
    private List<DifferentialStyle> ParseDifferentialFormats(string? content)
    {
        if (string.IsNullOrEmpty(content))
            return [];

        List<DifferentialStyle> styles = [];
        try
        {
            // ReadInnerXml re-declares in-scope namespaces on each child, so a bare wrapper
            // element is enough to parse what is otherwise a multi-root fragment.
            using var xml = XmlReader.Create(new StringReader("<dxfs>" + content + "</dxfs>"), XmlDefaults.ReaderSettings);
            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "dxf" && xml.Depth == 1)
                    styles.Add(ParseDifferentialFormat(xml));
            }
        }
        catch (XmlException)
        {
            // The raw markup is replayed on save regardless; an unparseable table only loses
            // the readable view.
            return [];
        }

        return styles;
    }

    private DifferentialStyle ParseDifferentialFormat(XmlReader xml)
    {
        var style = DifferentialStyle.Empty;
        if (xml.IsEmptyElement)
            return style;

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "dxf")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "font":
                    style = ParseDifferentialFont(xml, style);
                    break;
                case "numFmt" when xml.GetAttribute("formatCode") is { } code:
                    style = style with { NumberFormat = code };
                    break;
                case "fill":
                    // A dxf fill names its paint through bgColor and usually has no
                    // patternType at all; some writers use fgColor instead.
                    (ExcelColor foreground, ExcelColor background) = ParseFillColors(xml);
                    ExcelColor paint = !background.IsAuto ? background : foreground;
                    if (!paint.IsAuto)
                        style = style with { FillColor = paint };
                    break;
                case "border":
                    BorderStyle border = ParseBorder(xml);
                    if (border != BorderStyle.None)
                        style = style with { Border = border };
                    break;
                default:
                    break;
            }
        }

        return style;
    }

    /// <summary>Reads just the two colors out of a <c>fill</c> element, pattern or not.</summary>
    private static (ExcelColor Foreground, ExcelColor Background) ParseFillColors(XmlReader xml)
    {
        ExcelColor foreground = ExcelColor.Auto;
        ExcelColor background = ExcelColor.Auto;
        if (xml.IsEmptyElement)
            return (foreground, background);

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "fill")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "fgColor")
                foreground = ParseColor(xml);
            else if (xml.LocalName == "bgColor")
                background = ParseColor(xml);
        }

        return (foreground, background);
    }

    /// <summary>The font overrides of a <c>dxf</c>: only what the element mentions is set.</summary>
    private static DifferentialStyle ParseDifferentialFont(XmlReader xml, DifferentialStyle style)
    {
        if (xml.IsEmptyElement)
            return style;

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "font")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            style = xml.LocalName switch
            {
                "b" => style with { Bold = IsFlagSet(xml) },
                "i" => style with { Italic = IsFlagSet(xml) },
                "strike" => style with { Strikethrough = IsFlagSet(xml) },
                "u" => style with { Underline = xml.GetAttribute("val") != "none" },
                "color" => style with { FontColor = ParseColor(xml) },
                _ => style,
            };
        }

        return style;
    }
}
