using System.Globalization;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Generates the <c>conditionalFormatting</c> blocks of a worksheet from the model. Used only
/// once the rules have been edited; an untouched sheet gets its original markup replayed.
/// </summary>
internal static class ConditionalFormatWriter
{
    /// <summary>
    /// Writes the sheet's rules, one <c>conditionalFormatting</c> block per distinct set of
    /// ranges. Rules that share a range are grouped, which is how Excel writes them and what
    /// keeps <c>stopIfTrue</c> meaningful.
    /// </summary>
    /// <param name="xml">Sheet writer positioned where the blocks belong.</param>
    /// <param name="rules">Rules of the sheet, in priority order.</param>
    /// <param name="styleIds">Differential format index per rule, or -1 for a rule that paints nothing.</param>
    public static void Write(
        Utf8XmlWriter xml, IReadOnlyList<ConditionalFormatRule> rules, IReadOnlyList<int> styleIds)
    {
        int index = 0;
        while (index < rules.Count)
        {
            IReadOnlyList<RangeRef> ranges = rules[index].Ranges;
            int last = index;
            while (last + 1 < rules.Count && SameRanges(rules[last + 1].Ranges, ranges))
                last++;

            xml.WriteRaw("<conditionalFormatting sqref=\""u8);
            WriteRanges(xml, ranges);
            xml.WriteRaw("\">"u8);
            for (int i = index; i <= last; i++)
                WriteRule(xml, rules[i], styleIds[i]);
            xml.WriteRaw("</conditionalFormatting>"u8);
            index = last + 1;
        }
    }

    private static bool SameRanges(IReadOnlyList<RangeRef> left, IReadOnlyList<RangeRef> right)
    {
        if (left.Count != right.Count)
            return false;
        for (int i = 0; i < left.Count; i++)
        {
            if (left[i] != right[i])
                return false;
        }

        return true;
    }

    private static void WriteRanges(Utf8XmlWriter xml, IReadOnlyList<RangeRef> ranges)
    {
        for (int i = 0; i < ranges.Count; i++)
        {
            if (i > 0)
                xml.WriteRaw(" "u8);
            xml.WriteFormattable(ranges[i]);
        }
    }

    private static void WriteRule(Utf8XmlWriter xml, ConditionalFormatRule rule, int styleId)
    {
        if (rule.RawXml is { } raw)
        {
            RawXml.Write(xml, raw);
            return;
        }

        xml.WriteRaw("<cfRule type=\""u8);
        xml.WriteRaw(TypeName(rule.Type));
        if (styleId >= 0)
        {
            xml.WriteRaw("\" dxfId=\""u8);
            xml.WriteInt32(styleId);
        }

        xml.WriteRaw("\" priority=\""u8);
        xml.WriteInt32(rule.Priority);
        xml.WriteRaw("\""u8);
        WriteRuleAttributes(xml, rule);
        if (rule.StopIfTrue)
            xml.WriteRaw(" stopIfTrue=\"1\""u8);

        // Excel evaluates a text rule through a formula rather than the attributes, and a rule
        // built from the model has none, so the equivalent one is generated here.
        string? generated = rule.Formulas.Count == 0 && IsTextRule(rule.Type)
            ? TextRuleFormula(rule, rule.Ranges[0])
            : null;

        if (!HasChildren(rule) && generated is null)
        {
            xml.WriteRaw("/>"u8);
            return;
        }

        xml.WriteRaw(">"u8);
        foreach (string formula in rule.Formulas)
            WriteFormula(xml, formula);
        if (generated is not null)
            WriteFormula(xml, generated);

        WriteColorScale(xml, rule.ColorScale);
        WriteDataBar(xml, rule.DataBar);
        WriteIconSet(xml, rule.IconSet);
        xml.WriteRaw("</cfRule>"u8);
    }

    private static void WriteFormula(Utf8XmlWriter xml, string formula)
    {
        xml.WriteRaw("<formula>"u8);
        xml.WriteText(FutureFunctions.QualifyAll(formula)!);
        xml.WriteRaw("</formula>"u8);
    }

    private static bool IsTextRule(ConditionalFormatType type) => type is
        ConditionalFormatType.ContainsText or ConditionalFormatType.NotContainsText or
        ConditionalFormatType.BeginsWith or ConditionalFormatType.EndsWith;

    private static bool HasChildren(ConditionalFormatRule rule) =>
        rule.Formulas.Count > 0 || rule.ColorScale is not null || rule.DataBar is not null || rule.IconSet is not null;

    private static void WriteRuleAttributes(Utf8XmlWriter xml, ConditionalFormatRule rule)
    {
        switch (rule.Type)
        {
            case ConditionalFormatType.CellIs:
                xml.WriteRaw(" operator=\""u8);
                xml.WriteRaw(OperatorName(rule.Operator));
                xml.WriteRaw("\""u8);
                break;

            case ConditionalFormatType.ContainsText:
            case ConditionalFormatType.NotContainsText:
            case ConditionalFormatType.BeginsWith:
            case ConditionalFormatType.EndsWith:
                xml.WriteRaw(" operator=\""u8);
                xml.WriteRaw(TextOperatorName(rule.Type));
                xml.WriteRaw("\" text=\""u8);
                xml.WriteAttributeText(rule.Text ?? "");
                xml.WriteRaw("\""u8);
                break;

            case ConditionalFormatType.Top10:
                xml.WriteRaw(" rank=\""u8);
                xml.WriteInt32(rule.Rank);
                xml.WriteRaw("\""u8);
                if (rule.Bottom)
                    xml.WriteRaw(" bottom=\"1\""u8);
                if (rule.Percent)
                    xml.WriteRaw(" percent=\"1\""u8);
                break;

            case ConditionalFormatType.AboveAverage:
                if (!rule.Above)
                    xml.WriteRaw(" aboveAverage=\"0\""u8);
                if (rule.EqualAverage)
                    xml.WriteRaw(" equalAverage=\"1\""u8);
                if (rule.StandardDeviation != 0)
                {
                    xml.WriteRaw(" stdDev=\""u8);
                    xml.WriteInt32(rule.StandardDeviation);
                    xml.WriteRaw("\""u8);
                }

                break;

            case ConditionalFormatType.TimePeriod:
                xml.WriteRaw(" timePeriod=\""u8);
                xml.WriteAttributeText(rule.TimePeriod ?? "today");
                xml.WriteRaw("\""u8);
                break;

            default:
                break;
        }
    }

    private static void WriteColorScale(Utf8XmlWriter xml, ColorScale? scale)
    {
        if (scale is null)
            return;

        xml.WriteRaw("<colorScale>"u8);
        foreach (ConditionalValue value in scale.Values)
            WriteValue(xml, value);
        foreach (ExcelColor color in scale.Colors)
            WriteColor(xml, color);
        xml.WriteRaw("</colorScale>"u8);
    }

    private static void WriteDataBar(Utf8XmlWriter xml, DataBar? bar)
    {
        if (bar is null)
            return;

        xml.WriteRaw("<dataBar"u8);
        if (!bar.ShowValue)
            xml.WriteRaw(" showValue=\"0\""u8);
        xml.WriteRaw(">"u8);
        WriteValue(xml, bar.Minimum);
        WriteValue(xml, bar.Maximum);
        WriteColor(xml, bar.Color);
        xml.WriteRaw("</dataBar>"u8);
    }

    private static void WriteIconSet(Utf8XmlWriter xml, IconSet? icons)
    {
        if (icons is null)
            return;

        xml.WriteRaw("<iconSet iconSet=\""u8);
        xml.WriteAttributeText(icons.Name);
        xml.WriteRaw("\""u8);
        if (!icons.ShowValue)
            xml.WriteRaw(" showValue=\"0\""u8);
        if (icons.Reverse)
            xml.WriteRaw(" reverse=\"1\""u8);
        xml.WriteRaw(">"u8);
        foreach (ConditionalValue value in icons.Values)
            WriteValue(xml, value);
        xml.WriteRaw("</iconSet>"u8);
    }

    private static void WriteValue(Utf8XmlWriter xml, ConditionalValue value)
    {
        xml.WriteRaw("<cfvo type=\""u8);
        xml.WriteRaw(ValueTypeName(value.Type));
        if (value.Value is { Length: > 0 } text)
        {
            xml.WriteRaw("\" val=\""u8);
            xml.WriteAttributeText(text);
        }
        else if (value.Type is ConditionalValueType.Number or ConditionalValueType.Percent or ConditionalValueType.Percentile)
        {
            xml.WriteRaw("\" val=\"0"u8);
        }

        xml.WriteRaw("\"/>"u8);
    }

    private static void WriteColor(Utf8XmlWriter xml, ExcelColor color)
    {
        if (color.IsAuto)
        {
            xml.WriteRaw("<color auto=\"1\"/>"u8);
            return;
        }

        xml.WriteRaw("<color rgb=\""u8);
        xml.WriteAttributeText(color.ToHex());
        xml.WriteRaw("\"/>"u8);
    }

    private static ReadOnlySpan<byte> TypeName(ConditionalFormatType type) => type switch
    {
        ConditionalFormatType.CellIs => "cellIs"u8,
        ConditionalFormatType.Expression => "expression"u8,
        ConditionalFormatType.ColorScale => "colorScale"u8,
        ConditionalFormatType.DataBar => "dataBar"u8,
        ConditionalFormatType.IconSet => "iconSet"u8,
        ConditionalFormatType.Top10 => "top10"u8,
        ConditionalFormatType.AboveAverage => "aboveAverage"u8,
        ConditionalFormatType.ContainsText => "containsText"u8,
        ConditionalFormatType.NotContainsText => "notContainsText"u8,
        ConditionalFormatType.BeginsWith => "beginsWith"u8,
        ConditionalFormatType.EndsWith => "endsWith"u8,
        ConditionalFormatType.ContainsBlanks => "containsBlanks"u8,
        ConditionalFormatType.NotContainsBlanks => "notContainsBlanks"u8,
        ConditionalFormatType.ContainsErrors => "containsErrors"u8,
        ConditionalFormatType.NotContainsErrors => "notContainsErrors"u8,
        ConditionalFormatType.DuplicateValues => "duplicateValues"u8,
        ConditionalFormatType.UniqueValues => "uniqueValues"u8,
        _ => "timePeriod"u8,
    };

    private static ReadOnlySpan<byte> OperatorName(ConditionalOperator op) => op switch
    {
        ConditionalOperator.Between => "between"u8,
        ConditionalOperator.NotBetween => "notBetween"u8,
        ConditionalOperator.NotEqual => "notEqual"u8,
        ConditionalOperator.GreaterThan => "greaterThan"u8,
        ConditionalOperator.LessThan => "lessThan"u8,
        ConditionalOperator.GreaterThanOrEqual => "greaterThanOrEqual"u8,
        ConditionalOperator.LessThanOrEqual => "lessThanOrEqual"u8,
        _ => "equal"u8,
    };

    private static ReadOnlySpan<byte> TextOperatorName(ConditionalFormatType type) => type switch
    {
        ConditionalFormatType.NotContainsText => "notContains"u8,
        ConditionalFormatType.BeginsWith => "beginsWith"u8,
        ConditionalFormatType.EndsWith => "endsWith"u8,
        _ => "containsText"u8,
    };

    private static ReadOnlySpan<byte> ValueTypeName(ConditionalValueType type) => type switch
    {
        ConditionalValueType.Minimum => "min"u8,
        ConditionalValueType.Maximum => "max"u8,
        ConditionalValueType.Percent => "percent"u8,
        ConditionalValueType.Percentile => "percentile"u8,
        ConditionalValueType.Formula => "formula"u8,
        _ => "num"u8,
    };

    /// <summary>The formula that expresses a text rule, written for the range's top-left cell.</summary>
    private static string TextRuleFormula(ConditionalFormatRule rule, RangeRef first)
    {
        string cell = new CellRef(first.Start.Row, first.Start.Column).ToString();
        string needle = "\"" + (rule.Text ?? "").Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
        int length = (rule.Text ?? "").Length;
        string count = length.ToString(CultureInfo.InvariantCulture);
        return rule.Type switch
        {
            ConditionalFormatType.ContainsText => $"NOT(ISERROR(SEARCH({needle},{cell})))",
            ConditionalFormatType.NotContainsText => $"ISERROR(SEARCH({needle},{cell}))",
            ConditionalFormatType.BeginsWith => $"LEFT({cell},{count})={needle}",
            _ => $"RIGHT({cell},{count})={needle}",
        };
    }
}
