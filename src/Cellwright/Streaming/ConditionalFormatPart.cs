using System.Globalization;
using System.Xml;
using Cellwright.IO;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing the <c>conditionalFormatting</c> blocks of a worksheet. A rule the model
/// cannot represent — one carrying x14 extensions, say — is kept as its own XML and written
/// back unchanged, so editing the rules next to it costs nothing.
/// </summary>
internal static class ConditionalFormatPart
{
    private const string X14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string XmNamespace = "http://schemas.microsoft.com/office/excel/2006/main";

    /// <summary>
    /// Finds extension-only conditional-format rules and the worksheet-wide priorities they own.
    /// A no-priority x14 rule cannot own a second worksheet priority. This scan does not infer a
    /// companion relationship; the later data-bar stage must match one unambiguous GUID exactly.
    /// </summary>
    public static IReadOnlySet<int>? ReadExtensionPriorities(
        string? extensions, out bool hasConditionalFormats)
    {
        hasConditionalFormats = false;
        if (string.IsNullOrWhiteSpace(extensions))
            return null;

        HashSet<int>? priorities = null;
        try
        {
            using var xml = XmlReader.Create(
                new StringReader(extensions), XmlDefaults.ReaderSettings);
            while (xml.Read())
            {
                if (xml.NodeType != XmlNodeType.Element ||
                    !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal))
                {
                    continue;
                }

                if (xml.LocalName == "conditionalFormatting")
                {
                    hasConditionalFormats = true;
                    continue;
                }

                if (xml.LocalName != "cfRule")
                    continue;

                hasConditionalFormats = true;
                if (int.TryParse(
                    xml.GetAttribute("priority"),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out int priority) && priority > 0)
                {
                    (priorities ??= []).Add(priority);
                }
            }
        }
        catch (XmlException)
        {
            // extLst is preserved as opaque markup. A malformed optional extension must not make
            // an otherwise readable worksheet fail; any priorities found before the error remain
            // reserved conservatively.
        }

        return priorities;
    }

    /// <summary>
    /// Builds a read-only semantic view of extension-only conditional formats. The view is kept
    /// apart from the worksheet's editable collection: saving continues to replay the original
    /// <c>extLst</c>, so a supported rule is neither rewritten as a base rule nor emitted twice.
    /// </summary>
    public static ExtensionConditionalFormatProjection? ReadExtensionConditionalFormats(
        string? extensions,
        out IReadOnlySet<int>? priorities,
        out bool hasConditionalFormats)
    {
        priorities = ReadExtensionPriorities(extensions, out hasConditionalFormats);
        if (!hasConditionalFormats || string.IsNullOrWhiteSpace(extensions))
            return null;

        List<ConditionalFormatRule> rules = [];
        bool hasUnsupportedRules = false;
        bool sawBlock = false;
        try
        {
            using var xml = XmlReader.Create(
                new StringReader(extensions), XmlDefaults.ReaderSettings);
            bool advance = true;
            while (true)
            {
                if (advance && !xml.Read())
                    break;
                advance = true;

                if (xml.NodeType != XmlNodeType.Element ||
                    xml.LocalName != "conditionalFormatting" ||
                    !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal))
                {
                    continue;
                }

                sawBlock = true;
                string block = xml.ReadOuterXml();
                ReadExtensionConditionalFormatting(block, rules, ref hasUnsupportedRules);
                advance = false;
            }
        }
        catch (XmlException)
        {
            // The raw extension remains available for round-tripping. Complete rules parsed
            // before the malformed tail are safe to evaluate; the remainder is reported once.
            hasUnsupportedRules = true;
        }

        if (!sawBlock)
            hasUnsupportedRules = true;

        return new ExtensionConditionalFormatProjection(rules, hasUnsupportedRules);
    }

    /// <summary>Reads one x14 conditional-formatting block and its trailing xm:sqref.</summary>
    private static void ReadExtensionConditionalFormatting(
        string block,
        List<ConditionalFormatRule> projected,
        ref bool hasUnsupportedRules)
    {
        using var xml = XmlReader.Create(new StringReader(block), XmlDefaults.ReaderSettings);
        if (xml.MoveToContent() != XmlNodeType.Element ||
            xml.LocalName != "conditionalFormatting" ||
            !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal))
        {
            hasUnsupportedRules = true;
            return;
        }

        if (!TryBooleanAttribute(xml, "pivot", defaultValue: false, out bool pivot) || pivot)
        {
            hasUnsupportedRules = true;
            return;
        }

        int rootDepth = xml.Depth;
        List<string> rawRules = [];
        string? sqref = null;
        bool invalidBlock = xml.IsEmptyElement;
        bool advance = true;
        while (!invalidBlock)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.Depth == rootDepth)
                break;
            if (xml.NodeType != XmlNodeType.Element || xml.Depth != rootDepth + 1)
                continue;

            if (xml.LocalName == "cfRule" &&
                string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal))
            {
                rawRules.Add(xml.ReadOuterXml());
                advance = false;
            }
            else if (xml.LocalName == "sqref" &&
                     string.Equals(xml.NamespaceURI, XmNamespace, StringComparison.Ordinal) &&
                     sqref is null)
            {
                sqref = xml.ReadElementContentAsString();
                advance = false;
            }
            else
            {
                invalidBlock = true;
                xml.Skip();
                advance = false;
            }
        }

        if (invalidBlock || rawRules.Count == 0 || !TryParseExtensionRanges(sqref, out IReadOnlyList<RangeRef> ranges))
        {
            hasUnsupportedRules = true;
            return;
        }

        foreach (string rawRule in rawRules)
        {
            if (TryReadExtensionIconSetRule(rawRule, ranges, out ConditionalFormatRule rule))
                projected.Add(rule);
            else
                hasUnsupportedRules = true;
        }
    }

    /// <summary>Reads the supported, independent x14 icon-set rule subset.</summary>
    private static bool TryReadExtensionIconSetRule(
        string raw,
        IReadOnlyList<RangeRef> ranges,
        out ConditionalFormatRule rule)
    {
        rule = null!;
        using var xml = XmlReader.Create(new StringReader(raw), XmlDefaults.ReaderSettings);
        if (xml.MoveToContent() != XmlNodeType.Element ||
            xml.LocalName != "cfRule" ||
            !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal) ||
            xml.GetAttribute("type") != "iconSet" ||
            !int.TryParse(
                xml.GetAttribute("priority"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out int priority) ||
            priority <= 0 ||
            !TryBooleanAttribute(xml, "stopIfTrue", defaultValue: false, out bool stopIfTrue))
        {
            return false;
        }

        int rootDepth = xml.Depth;
        List<string> formulas = [];
        IconSet? iconSet = null;
        bool invalid = xml.IsEmptyElement;
        bool advance = true;
        while (!invalid)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.Depth == rootDepth)
                break;
            if (xml.NodeType != XmlNodeType.Element || xml.Depth != rootDepth + 1)
                continue;

            if (xml.LocalName == "f" &&
                string.Equals(xml.NamespaceURI, XmNamespace, StringComparison.Ordinal) &&
                formulas.Count == 0)
            {
                string formula = xml.ReadElementContentAsString();
                if (string.IsNullOrWhiteSpace(formula) || formula.Length > SheetLimits.MaxCellTextLength)
                    invalid = true;
                else
                    formulas.Add(FutureFunctions.UnqualifyAll(formula)!);
                advance = false;
            }
            else if (xml.LocalName == "iconSet" &&
                     string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal) &&
                     iconSet is null)
            {
                string rawIconSet = xml.ReadOuterXml();
                if (!TryReadExtensionIconSet(rawIconSet, out iconSet))
                    invalid = true;
                advance = false;
            }
            else
            {
                invalid = true;
                xml.Skip();
                advance = false;
            }
        }

        if (invalid || iconSet is null)
            return false;

        rule = new ConditionalFormatRule
        {
            Ranges = ranges,
            Type = ConditionalFormatType.IconSet,
            Formulas = formulas,
            Priority = priority,
            StopIfTrue = stopIfTrue,
            IconSet = iconSet,
        };
        return true;
    }

    /// <summary>Reads a non-custom x14 icon set and validates its schema cardinality.</summary>
    private static bool TryReadExtensionIconSet(string raw, out IconSet? iconSet)
    {
        iconSet = null;
        using var xml = XmlReader.Create(new StringReader(raw), XmlDefaults.ReaderSettings);
        if (xml.MoveToContent() != XmlNodeType.Element ||
            xml.LocalName != "iconSet" ||
            !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal) ||
            !TryBooleanAttribute(xml, "custom", defaultValue: false, out bool custom) ||
            custom ||
            !TryBooleanAttribute(xml, "showValue", defaultValue: true, out bool showValue) ||
            !TryBooleanAttribute(xml, "reverse", defaultValue: false, out bool reverse))
        {
            return false;
        }

        string name = xml.GetAttribute("iconSet") ?? "3TrafficLights1";
        if (!TryIconSetSize(name, out int expectedCount))
            return false;

        int rootDepth = xml.Depth;
        List<ConditionalValue> values = [];
        bool invalid = xml.IsEmptyElement;
        bool advance = true;
        while (!invalid)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.Depth == rootDepth)
                break;
            if (xml.NodeType != XmlNodeType.Element || xml.Depth != rootDepth + 1)
                continue;

            if (xml.LocalName == "cfvo" &&
                string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal))
            {
                string rawValue = xml.ReadOuterXml();
                if (TryReadExtensionValue(rawValue, out ConditionalValue value))
                    values.Add(value);
                else
                    invalid = true;
                advance = false;
            }
            else
            {
                // cfIcon belongs to custom sets; extensions and unknown children may change the
                // meaning of the set, so the conservative projection does not half-apply them.
                invalid = true;
                xml.Skip();
                advance = false;
            }
        }

        if (invalid || values.Count is < 3 or > 5 || expectedCount != 0 && values.Count != expectedCount)
            return false;

        iconSet = new IconSet
        {
            Name = name,
            Values = values,
            ShowValue = showValue,
            Reverse = reverse,
        };
        return true;
    }

    /// <summary>Reads one x14 cfvo whose value is carried by its xm:f child.</summary>
    private static bool TryReadExtensionValue(string raw, out ConditionalValue value)
    {
        value = default;
        using var xml = XmlReader.Create(new StringReader(raw), XmlDefaults.ReaderSettings);
        if (xml.MoveToContent() != XmlNodeType.Element ||
            xml.LocalName != "cfvo" ||
            !string.Equals(xml.NamespaceURI, X14Namespace, StringComparison.Ordinal) ||
            !TryValueType(xml.GetAttribute("type"), out ConditionalValueType type) ||
            !TryBooleanAttribute(xml, "gte", defaultValue: true, out bool greaterThanOrEqual))
        {
            return false;
        }

        int rootDepth = xml.Depth;
        string? text = null;
        bool hasChildren = !xml.IsEmptyElement;
        bool invalid = false;
        bool advance = true;
        while (hasChildren && !invalid)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.Depth == rootDepth)
                break;
            if (xml.NodeType != XmlNodeType.Element || xml.Depth != rootDepth + 1)
                continue;

            if (xml.LocalName == "f" &&
                string.Equals(xml.NamespaceURI, XmNamespace, StringComparison.Ordinal) &&
                text is null)
            {
                text = xml.ReadElementContentAsString().Trim();
                advance = false;
            }
            else
            {
                invalid = true;
                xml.Skip();
                advance = false;
            }
        }

        if (invalid)
        {
            return false;
        }

        // CT_Cfvo defines a missing X as zero, including for a formula-typed stop. Icon sets do
        // not permit min/max or autoMin/autoMax; those belong to the later data-bar projection.
        if (string.IsNullOrEmpty(text))
            text = "0";

        if (text.Length > SheetLimits.MaxCellTextLength || !IsValidExtensionValue(type, text))
        {
            return false;
        }

        value = new ConditionalValue(
            type,
            type == ConditionalValueType.Formula ? FutureFunctions.UnqualifyAll(text) : text)
        {
            GreaterThanOrEqual = greaterThanOrEqual,
        };
        return true;
    }

    private static bool TryParseExtensionRanges(string? sqref, out IReadOnlyList<RangeRef> ranges)
    {
        ranges = [];
        if (string.IsNullOrWhiteSpace(sqref))
            return false;

        string[] references = sqref.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (references.Length == 0)
            return false;

        List<RangeRef> parsed = new(references.Length);
        foreach (string reference in references)
        {
            if (!RangeRef.TryParse(reference, null, out RangeRef range))
                return false;
            parsed.Add(range);
        }

        ranges = parsed;
        return true;
    }

    private static bool TryBooleanAttribute(XmlReader xml, string name, bool defaultValue, out bool value)
    {
        if (xml.GetAttribute(name) is not { } text)
        {
            value = defaultValue;
            return true;
        }

        switch (text.Trim())
        {
            case "1" or "true":
                value = true;
                return true;
            case "0" or "false":
                value = false;
                return true;
            default:
                value = default;
                return false;
        }
    }

    private static bool TryValueType(string? text, out ConditionalValueType type)
    {
        type = text switch
        {
            "num" => ConditionalValueType.Number,
            "percent" => ConditionalValueType.Percent,
            "percentile" => ConditionalValueType.Percentile,
            "formula" => ConditionalValueType.Formula,
            _ => (ConditionalValueType)(-1),
        };
        return (int)type >= 0;
    }

    private static bool IsValidExtensionValue(ConditionalValueType type, string text)
    {
        if (type == ConditionalValueType.Formula)
            return !string.IsNullOrWhiteSpace(text);

        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            // SpreadsheetML permits a formula in xm:f even when the stop type is numeric,
            // percent or percentile. The evaluator will report a missing formula engine if one
            // is needed; rejecting it here would turn a supported rule opaque unnecessarily.
            return !string.IsNullOrWhiteSpace(text);
        }

        return double.IsFinite(number);
    }

    /// <summary>Returns the fixed cardinality of the x14-only sets this projection supports.</summary>
    private static bool TryIconSetSize(string name, out int count)
    {
        count = name switch
        {
            "3Stars" or "3Triangles" => 3,
            "5Boxes" => 5,
            _ => -1,
        };
        return count >= 0;
    }

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
        int minimumLength = ParseInt(xml.GetAttribute("minLength"), 10);
        int maximumLength = ParseInt(xml.GetAttribute("maxLength"), 90);
        (List<ConditionalValue> values, List<ExcelColor> colors) = ReadStops(xml, "dataBar");
        if (values.Count < 2)
            return null;

        return new DataBar
        {
            Minimum = values[0],
            Maximum = values[1],
            Color = colors.Count > 0 ? colors[0] : ExcelColor.FromRgb(0x63, 0x8E, 0xC6),
            MinimumLength = minimumLength,
            MaximumLength = maximumLength,
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
                    values.Add(new ConditionalValue(
                        ParseValueType(xml.GetAttribute("type")),
                        xml.GetAttribute("val"))
                    {
                        GreaterThanOrEqual = container != "iconSet" ||
                            xml.GetAttribute("gte") is not { } gte || IsTrue(gte),
                    });
                    break;
                case "color":
                    colors.Add(StylesPartParser.ParseColor(xml));
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
        int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : fallback;

    private static bool IsTrue(string? value) => value is "1" or "true" or "TRUE";
}
