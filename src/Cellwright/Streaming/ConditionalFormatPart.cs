using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing the <c>conditionalFormatting</c> blocks of a worksheet. A rule the model
/// cannot represent — one carrying x14 extensions, say — is kept as its own XML and written
/// back unchanged, so editing the rules next to it costs nothing.
/// </summary>
internal static class ConditionalFormatPart
{
    /// <summary>Parses one captured <c>conditionalFormatting</c> element into rules.</summary>
    public static void Read(string block, ISheetContentSink sink)
    {
        using var xml = XmlReader.Create(new StringReader(block), XmlDefaults.ReaderSettings);
        IReadOnlyList<RangeRef>? ranges = null;
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "conditionalFormatting")
            {
                ranges = TryParseRanges(xml.GetAttribute("sqref"));
                continue;
            }

            if (xml.LocalName != "cfRule" || ranges is null)
                continue;

            sink.OnConditionalFormat(ReadRule(xml, ranges));
            advance = false;
        }
    }

    private static IReadOnlyList<RangeRef>? TryParseRanges(string? sqref)
    {
        if (string.IsNullOrWhiteSpace(sqref))
            return null;

        try
        {
            return DataValidation.ParseRanges(sqref);
        }
        catch (FormatException)
        {
            return null;
        }
    }

    /// <summary>
    /// Reads one rule. The whole element is taken as text first: everything the attributes and
    /// the known children do not explain has to travel with the rule, and there is no second
    /// pass over a forward-only reader.
    /// </summary>
    private static ConditionalFormatRule ReadRule(XmlReader xml, IReadOnlyList<RangeRef> ranges)
    {
        ConditionalFormatType? type = ParseType(xml.GetAttribute("type"));
        int priority = ParseInt(xml.GetAttribute("priority"), 1);
        int? styleId = int.TryParse(xml.GetAttribute("dxfId"), NumberStyles.None, CultureInfo.InvariantCulture, out int dxf)
            ? dxf
            : null;
        bool stopIfTrue = IsTrue(xml.GetAttribute("stopIfTrue"));
        ConditionalOperator op = ParseOperator(xml.GetAttribute("operator"));
        string? text = xml.GetAttribute("text");
        string? timePeriod = xml.GetAttribute("timePeriod");
        int rank = ParseInt(xml.GetAttribute("rank"), 10);
        bool bottom = IsTrue(xml.GetAttribute("bottom"));
        bool percent = IsTrue(xml.GetAttribute("percent"));
        bool above = xml.GetAttribute("aboveAverage") is not { } value || IsTrue(value);
        bool equalAverage = IsTrue(xml.GetAttribute("equalAverage"));
        int standardDeviation = ParseInt(xml.GetAttribute("stdDev"), 0);
        string raw = xml.ReadOuterXml();

        var opaque = new ConditionalFormatRule
        {
            Ranges = ranges,
            Priority = priority,
            StopIfTrue = stopIfTrue,
            SourceStyleId = styleId,
            RawXml = raw,
        };

        if (type is not { } known)
            return opaque;

        List<string> formulas = [];
        ColorScale? scale = null;
        DataBar? bar = null;
        IconSet? icons = null;
        using var fragment = XmlReader.Create(new StringReader(raw), XmlDefaults.ReaderSettings);
        bool advance = true;
        while (true)
        {
            if (advance && !fragment.Read())
                break;
            advance = true;

            if (fragment.NodeType != XmlNodeType.Element)
                continue;

            switch (fragment.LocalName)
            {
                // Reading the content leaves the reader on the node after the element, so the
                // next bound already has one; two formulas in a row would otherwise lose one.
                case "formula":
                    formulas.Add(FutureFunctions.UnqualifyAll(fragment.ReadElementContentAsString())!);
                    advance = false;
                    break;
                case "colorScale":
                    scale = ReadColorScale(fragment);
                    break;
                case "dataBar":
                    bar = ReadDataBar(fragment);
                    break;
                case "iconSet":
                    icons = ReadIconSet(fragment);
                    break;

                // An extension list holds the modern rule variants, which only Excel's own
                // x14 namespace describes. Such a rule stays opaque rather than half-read.
                case "extLst":
                    return opaque;
                default:
                    break;
            }
        }

        return new ConditionalFormatRule
        {
            Ranges = ranges,
            Type = known,
            Operator = op,
            Formulas = formulas,
            Text = text,
            Rank = rank,
            Bottom = bottom,
            Percent = percent,
            Above = above,
            EqualAverage = equalAverage,
            StandardDeviation = standardDeviation,
            TimePeriod = timePeriod,
            Priority = priority,
            StopIfTrue = stopIfTrue,
            ColorScale = scale,
            DataBar = bar,
            IconSet = icons,
            SourceStyleId = styleId,
        };
    }

    private static ColorScale? ReadColorScale(XmlReader xml)
    {
        (List<ConditionalValue> values, List<ExcelColor> colors) = ReadStops(xml, "colorScale");
        return values.Count >= 2 && colors.Count == values.Count
            ? new ColorScale { Values = values, Colors = colors }
            : null;
    }

    private static DataBar? ReadDataBar(XmlReader xml)
    {
        bool showValue = xml.GetAttribute("showValue") is not { } shown || IsTrue(shown);
        (List<ConditionalValue> values, List<ExcelColor> colors) = ReadStops(xml, "dataBar");
        if (values.Count < 2)
            return null;

        return new DataBar
        {
            Minimum = values[0],
            Maximum = values[1],
            Color = colors.Count > 0 ? colors[0] : ExcelColor.FromRgb(0x63, 0x8E, 0xC6),
            ShowValue = showValue,
        };
    }

    private static IconSet? ReadIconSet(XmlReader xml)
    {
        string name = xml.GetAttribute("iconSet") ?? "3TrafficLights1";
        bool showValue = xml.GetAttribute("showValue") is not { } shown || IsTrue(shown);
        bool reverse = IsTrue(xml.GetAttribute("reverse"));
        (List<ConditionalValue> values, _) = ReadStops(xml, "iconSet");
        return values.Count == 0
            ? null
            : new IconSet { Name = name, Values = values, ShowValue = showValue, Reverse = reverse };
    }

    /// <summary>Reads the <c>cfvo</c> stops and <c>color</c> elements of a graded rule.</summary>
    private static (List<ConditionalValue> Values, List<ExcelColor> Colors) ReadStops(XmlReader xml, string container)
    {
        List<ConditionalValue> values = [];
        List<ExcelColor> colors = [];
        if (xml.IsEmptyElement)
            return (values, colors);

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == container)
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "cfvo":
                    values.Add(new ConditionalValue(ParseValueType(xml.GetAttribute("type")), xml.GetAttribute("val")));
                    break;
                case "color":
                    colors.Add(ExcelColor.TryParseHex(xml.GetAttribute("rgb"), out ExcelColor color) ? color : ExcelColor.Auto);
                    break;
                default:
                    break;
            }
        }

        return (values, colors);
    }

    private static ConditionalFormatType? ParseType(string? type) => type switch
    {
        "cellIs" => ConditionalFormatType.CellIs,
        "expression" => ConditionalFormatType.Expression,
        "colorScale" => ConditionalFormatType.ColorScale,
        "dataBar" => ConditionalFormatType.DataBar,
        "iconSet" => ConditionalFormatType.IconSet,
        "top10" => ConditionalFormatType.Top10,
        "aboveAverage" => ConditionalFormatType.AboveAverage,
        "containsText" => ConditionalFormatType.ContainsText,
        "notContainsText" => ConditionalFormatType.NotContainsText,
        "beginsWith" => ConditionalFormatType.BeginsWith,
        "endsWith" => ConditionalFormatType.EndsWith,
        "containsBlanks" => ConditionalFormatType.ContainsBlanks,
        "notContainsBlanks" => ConditionalFormatType.NotContainsBlanks,
        "containsErrors" => ConditionalFormatType.ContainsErrors,
        "notContainsErrors" => ConditionalFormatType.NotContainsErrors,
        "duplicateValues" => ConditionalFormatType.DuplicateValues,
        "uniqueValues" => ConditionalFormatType.UniqueValues,
        "timePeriod" => ConditionalFormatType.TimePeriod,
        _ => null,
    };

    private static ConditionalOperator ParseOperator(string? text) => text switch
    {
        "between" => ConditionalOperator.Between,
        "notBetween" => ConditionalOperator.NotBetween,
        "equal" => ConditionalOperator.Equal,
        "notEqual" => ConditionalOperator.NotEqual,
        "greaterThan" => ConditionalOperator.GreaterThan,
        "lessThan" => ConditionalOperator.LessThan,
        "greaterThanOrEqual" => ConditionalOperator.GreaterThanOrEqual,
        "lessThanOrEqual" => ConditionalOperator.LessThanOrEqual,
        _ => ConditionalOperator.None,
    };

    private static ConditionalValueType ParseValueType(string? text) => text switch
    {
        "min" => ConditionalValueType.Minimum,
        "max" => ConditionalValueType.Maximum,
        "percent" => ConditionalValueType.Percent,
        "percentile" => ConditionalValueType.Percentile,
        "formula" => ConditionalValueType.Formula,
        _ => ConditionalValueType.Number,
    };

    private static int ParseInt(string? text, int fallback) =>
        int.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int value) ? value : fallback;

    private static bool IsTrue(string? value) => value is "1" or "true" or "TRUE";
}
