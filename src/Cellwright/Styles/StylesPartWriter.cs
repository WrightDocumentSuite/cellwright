using Cellwright.Xml;

namespace Cellwright.Styles;

/// <summary>
/// Serializes interned <see cref="CellStyle"/> records into the styles part:
/// components (fonts, fills, borders, custom number formats) are deduplicated
/// and referenced by the cell formats.
/// </summary>
internal static class StylesPartWriter
{
    /// <summary>Writes the complete styles XML for the given cell formats.</summary>
    public static void Write(
        Utf8XmlWriter xml,
        IReadOnlyList<CellStyle> styles,
        StylesMarkup? markup = null,
        IReadOnlyList<DifferentialStyle>? differentialStyles = null)
    {
        var fonts = new Dictionary<FontStyle, int> { [FontStyle.Default] = 0 };
        var fills = new Dictionary<FillStyle, int> { [FillStyle.None] = 0 };
        var borders = new Dictionary<BorderStyle, int> { [BorderStyle.None] = 0 };
        var customFormats = new Dictionary<string, int>(StringComparer.Ordinal);
        var formatIds = new int[styles.Count];

        for (int i = 0; i < styles.Count; i++)
        {
            CellStyle style = styles[i];
            fonts.TryAdd(style.Font, fonts.Count);
            fills.TryAdd(style.Fill, fills.Count);
            borders.TryAdd(style.Border, borders.Count);
            formatIds[i] = ResolveNumberFormatId(style.NumberFormat, customFormats);
        }

        // The masters of cellStyleXfs share the component tables with the cell formats: they
        // are parsed at load exactly so that the ids they reference stay true after the tables
        // are rebuilt.
        IReadOnlyList<CellStyle> masters = markup?.StyleFormats ?? [CellStyle.Default];
        var masterFormatIds = new int[masters.Count];
        for (int i = 0; i < masters.Count; i++)
        {
            CellStyle master = masters[i];
            fonts.TryAdd(master.Font, fonts.Count);
            fills.TryAdd(master.Fill, fills.Count);
            borders.TryAdd(master.Border, borders.Count);
            masterFormatIds[i] = ResolveNumberFormatId(master.NumberFormat, customFormats);
        }

        xml.WriteDeclaration();
        xml.WriteRaw("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""u8);
        RawXml.WriteAttributes(xml, markup?.RootAttributes);
        xml.WriteRaw(">"u8);
        WriteNumberFormats(xml, customFormats);
        WriteFonts(xml, OrderedKeys(fonts));
        WriteFills(xml, OrderedKeys(fills));
        WriteBorders(xml, OrderedKeys(borders));

        xml.WriteRaw("<cellStyleXfs count=\""u8);
        xml.WriteInt32(masters.Count);
        xml.WriteRaw("\">"u8);
        for (int i = 0; i < masters.Count; i++)
        {
            CellStyle master = masters[i];
            WriteCellFormat(
                xml, master, masterFormatIds[i], fonts[master.Font], fills[master.Fill], borders[master.Border],
                styleXfId: null);
        }

        xml.WriteRaw("</cellStyleXfs>"u8);

        xml.WriteRaw("<cellXfs count=\""u8);
        xml.WriteInt32(styles.Count);
        xml.WriteRaw("\">"u8);
        for (int i = 0; i < styles.Count; i++)
        {
            CellStyle style = styles[i];
            int styleXfId = (uint)style.StyleXfId < (uint)masters.Count ? style.StyleXfId : 0;
            WriteCellFormat(xml, style, formatIds[i], fonts[style.Font], fills[style.Fill], borders[style.Border], styleXfId);
        }

        xml.WriteRaw("</cellXfs>"u8);
        if (markup?.CellStyleEntries is { } cellStyles)
            RawXml.Write(xml, cellStyles);
        else
            xml.WriteRaw("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"u8);
        WriteDifferentialFormats(xml, markup, differentialStyles);
        RawXml.Write(xml, markup?.TableStyles);
        RawXml.Write(xml, markup?.Colors);
        RawXml.Write(xml, markup?.Extensions);
        xml.WriteRaw("</styleSheet>"u8);
    }

    private static List<TKey> OrderedKeys<TKey>(Dictionary<TKey, int> table) where TKey : notnull =>
        [.. table.OrderBy(static pair => pair.Value).Select(static pair => pair.Key)];

    /// <summary>
    /// Writes the differential formats: the ones the file already had, unchanged and in their
    /// original positions, followed by the ones conditional formatting asked for.
    /// </summary>
    private static void WriteDifferentialFormats(
        Utf8XmlWriter xml, StylesMarkup? markup, IReadOnlyList<DifferentialStyle>? added)
    {
        int preserved = markup?.DifferentialFormatCount ?? 0;
        int total = preserved + (added?.Count ?? 0);
        if (total == 0)
            return;

        xml.WriteRaw("<dxfs count=\""u8);
        xml.WriteInt32(total);
        xml.WriteRaw("\">"u8);
        RawXml.Write(xml, markup?.DifferentialFormats);
        if (added is not null)
        {
            foreach (DifferentialStyle style in added)
                WriteDifferentialFormat(xml, style);
        }

        xml.WriteRaw("</dxfs>"u8);
    }

    /// <summary>
    /// One <c>dxf</c>. Only the properties the format sets are written: a differential format
    /// that mentioned a font size would impose it on every cell the rule paints.
    /// </summary>
    private static void WriteDifferentialFormat(Utf8XmlWriter xml, DifferentialStyle style)
    {
        xml.WriteRaw("<dxf>"u8);
        if (style.Bold is not null || style.Italic is not null || style.Underline is not null ||
            style.Strikethrough is not null || !style.FontColor.IsAuto)
        {
            xml.WriteRaw("<font>"u8);
            WriteToggle(xml, "b"u8, style.Bold);
            WriteToggle(xml, "i"u8, style.Italic);
            WriteToggle(xml, "strike"u8, style.Strikethrough);
            WriteToggle(xml, "u"u8, style.Underline);
            if (!style.FontColor.IsAuto)
                WriteColorElement(xml, "<color rgb=\""u8, style.FontColor);
            xml.WriteRaw("</font>"u8);
        }

        if (style.NumberFormat is { Length: > 0 } code)
        {
            xml.WriteRaw("<numFmt numFmtId=\""u8);
            xml.WriteInt32(BuiltinNumberFormats.GetId(code) ?? BuiltinNumberFormats.FirstCustomId);
            xml.WriteRaw("\" formatCode=\""u8);
            xml.WriteAttributeText(code);
            xml.WriteRaw("\"/>"u8);
        }

        if (!style.FillColor.IsAuto)
        {
            // A dxf fill states the background through bgColor; Excel ignores fgColor here.
            xml.WriteRaw("<fill><patternFill><bgColor rgb=\""u8);
            xml.WriteAttributeText(style.FillColor.ToHex());
            xml.WriteRaw("\"/></patternFill></fill>"u8);
        }

        if (style.Border is { } border)
        {
            xml.WriteRaw("<border>"u8);
            WriteEdge(xml, "left"u8, border.Left);
            WriteEdge(xml, "right"u8, border.Right);
            WriteEdge(xml, "top"u8, border.Top);
            WriteEdge(xml, "bottom"u8, border.Bottom);
            xml.WriteRaw("</border>"u8);
        }

        xml.WriteRaw("</dxf>"u8);
    }

    private static void WriteToggle(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> name, bool? value)
    {
        if (value is null)
            return;

        xml.WriteRaw("<"u8);
        xml.WriteRaw(name);
        if (value == false)
            xml.WriteRaw(" val=\"0\""u8);
        xml.WriteRaw("/>"u8);
    }

    private static int ResolveNumberFormatId(string code, Dictionary<string, int> customFormats)
    {
        if (code.Length == 0)
            return 0;
        if (BuiltinNumberFormats.GetId(code) is { } builtin)
            return builtin;
        if (!customFormats.TryGetValue(code, out int id))
            customFormats[code] = id = BuiltinNumberFormats.FirstCustomId + customFormats.Count;
        return id;
    }

    private static void WriteNumberFormats(Utf8XmlWriter xml, Dictionary<string, int> customFormats)
    {
        if (customFormats.Count == 0)
            return;

        xml.WriteRaw("<numFmts count=\""u8);
        xml.WriteInt32(customFormats.Count);
        xml.WriteRaw("\">"u8);
        foreach ((string code, int id) in customFormats)
        {
            xml.WriteRaw("<numFmt numFmtId=\""u8);
            xml.WriteInt32(id);
            xml.WriteRaw("\" formatCode=\""u8);
            xml.WriteAttributeText(code);
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</numFmts>"u8);
    }

    private static void WriteFonts(Utf8XmlWriter xml, IReadOnlyCollection<FontStyle> fonts)
    {
        xml.WriteRaw("<fonts count=\""u8);
        xml.WriteInt32(fonts.Count);
        xml.WriteRaw("\">"u8);
        foreach (FontStyle font in fonts)
        {
            xml.WriteRaw("<font>"u8);
            FontMarkup.WriteChildren(xml, font, runProperties: false);
            xml.WriteRaw("</font>"u8);
        }

        xml.WriteRaw("</fonts>"u8);
    }

    private static void WriteFills(Utf8XmlWriter xml, IReadOnlyCollection<FillStyle> fills)
    {
        // Index 0 must be "none" and index 1 "gray125" by longstanding convention,
        // so real fills are written after the two reserved entries.
        xml.WriteRaw("<fills count=\""u8);
        xml.WriteInt32(fills.Count + 1);
        xml.WriteRaw("\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"u8);
        foreach (FillStyle fill in fills)
        {
            if (fill.IsEmpty)
                continue;

            if (fill.Gradient is { } gradient)
            {
                WriteGradientFill(xml, gradient);
                continue;
            }

            xml.WriteRaw("<fill><patternFill patternType=\""u8);
            xml.WriteRaw(PatternName(fill.Pattern));
            xml.WriteRaw("\">"u8);
            if (!fill.Color.IsAuto)
                WriteColorElement(xml, "<fgColor rgb=\""u8, fill.Color);

            // Excel writes indexed 64 — the system window color — where a solid fill has no
            // background of its own, and refuses to render some patterns without one.
            if (fill.BackgroundColor.IsAuto)
                xml.WriteRaw("<bgColor indexed=\"64\"/>"u8);
            else
                WriteColorElement(xml, "<bgColor rgb=\""u8, fill.BackgroundColor);

            xml.WriteRaw("</patternFill></fill>"u8);
        }

        xml.WriteRaw("</fills>"u8);
    }

    /// <summary>A <c>gradientFill</c> (§18.8.24): the sweep, then a <c>stop</c> per color.</summary>
    private static void WriteGradientFill(Utf8XmlWriter xml, GradientFill gradient)
    {
        xml.WriteRaw("<fill><gradientFill"u8);
        if (gradient.Type == GradientFillType.Path)
            xml.WriteRaw(" type=\"path\""u8);
        WriteGradientNumber(xml, " degree=\""u8, gradient.Degree);
        WriteGradientNumber(xml, " left=\""u8, gradient.Left);
        WriteGradientNumber(xml, " right=\""u8, gradient.Right);
        WriteGradientNumber(xml, " top=\""u8, gradient.Top);
        WriteGradientNumber(xml, " bottom=\""u8, gradient.Bottom);
        xml.WriteRaw(">"u8);
        foreach (GradientStop stop in gradient.Stops)
        {
            xml.WriteRaw("<stop position=\""u8);
            xml.WriteDouble(stop.Position);
            xml.WriteRaw("\">"u8);
            WriteColorElement(xml, "<color rgb=\""u8, stop.Color);
            xml.WriteRaw("</stop>"u8);
        }

        xml.WriteRaw("</gradientFill></fill>"u8);
    }

    private static void WriteGradientNumber(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> open, double value)
    {
        if (value == 0)
            return;

        xml.WriteRaw(open);
        xml.WriteDouble(value);
        xml.WriteRaw("\""u8);
    }

    private static ReadOnlySpan<byte> PatternName(FillPattern pattern) => pattern switch
    {
        FillPattern.Solid => "solid"u8,
        FillPattern.DarkGray => "darkGray"u8,
        FillPattern.MediumGray => "mediumGray"u8,
        FillPattern.LightGray => "lightGray"u8,
        FillPattern.Gray125 => "gray125"u8,
        FillPattern.Gray0625 => "gray0625"u8,
        FillPattern.DarkHorizontal => "darkHorizontal"u8,
        FillPattern.DarkVertical => "darkVertical"u8,
        FillPattern.DarkDown => "darkDown"u8,
        FillPattern.DarkUp => "darkUp"u8,
        FillPattern.DarkGrid => "darkGrid"u8,
        FillPattern.DarkTrellis => "darkTrellis"u8,
        FillPattern.LightHorizontal => "lightHorizontal"u8,
        FillPattern.LightVertical => "lightVertical"u8,
        FillPattern.LightDown => "lightDown"u8,
        FillPattern.LightUp => "lightUp"u8,
        FillPattern.LightGrid => "lightGrid"u8,
        _ => "lightTrellis"u8,
    };

    internal static void WriteUnderline(Utf8XmlWriter xml, UnderlineStyle underline)
    {
        switch (underline)
        {
            case UnderlineStyle.None:
                return;
            case UnderlineStyle.Single:
                xml.WriteRaw("<u/>"u8);
                return;
            default:
                xml.WriteRaw("<u val=\""u8);
                xml.WriteRaw(underline switch
                {
                    UnderlineStyle.Double => "double"u8,
                    UnderlineStyle.SingleAccounting => "singleAccounting"u8,
                    _ => "doubleAccounting"u8,
                });
                xml.WriteRaw("\"/>"u8);
                return;
        }
    }

    internal static void WriteIntElement(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> open, int value, bool write)
    {
        if (!write)
            return;

        xml.WriteRaw(open);
        xml.WriteInt32(value);
        xml.WriteRaw("\"/>"u8);
    }

    private static void WriteBorders(Utf8XmlWriter xml, IReadOnlyCollection<BorderStyle> borders)
    {
        xml.WriteRaw("<borders count=\""u8);
        xml.WriteInt32(borders.Count);
        xml.WriteRaw("\">"u8);
        foreach (BorderStyle border in borders)
        {
            xml.WriteRaw("<border"u8);
            if (border.DiagonalUp)
                xml.WriteRaw(" diagonalUp=\"1\""u8);
            if (border.DiagonalDown)
                xml.WriteRaw(" diagonalDown=\"1\""u8);
            xml.WriteRaw(">"u8);
            WriteEdge(xml, "left"u8, border.Left);
            WriteEdge(xml, "right"u8, border.Right);
            WriteEdge(xml, "top"u8, border.Top);
            WriteEdge(xml, "bottom"u8, border.Bottom);
            WriteEdge(xml, "diagonal"u8, border.Diagonal);
            xml.WriteRaw("</border>"u8);
        }

        xml.WriteRaw("</borders>"u8);
    }

    private static void WriteEdge(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> name, BorderEdge edge)
    {
        xml.WriteRaw("<"u8);
        xml.WriteRaw(name);
        if (edge.Line == BorderLineStyle.None)
        {
            xml.WriteRaw("/>"u8);
            return;
        }

        xml.WriteRaw(" style=\""u8);
        xml.WriteRaw(edge.Line switch
        {
            BorderLineStyle.Thin => "thin"u8,
            BorderLineStyle.Medium => "medium"u8,
            BorderLineStyle.Thick => "thick"u8,
            BorderLineStyle.Dashed => "dashed"u8,
            BorderLineStyle.Dotted => "dotted"u8,
            BorderLineStyle.Double => "double"u8,
            BorderLineStyle.MediumDashed => "mediumDashed"u8,
            BorderLineStyle.DashDot => "dashDot"u8,
            BorderLineStyle.MediumDashDot => "mediumDashDot"u8,
            BorderLineStyle.DashDotDot => "dashDotDot"u8,
            BorderLineStyle.MediumDashDotDot => "mediumDashDotDot"u8,
            BorderLineStyle.SlantDashDot => "slantDashDot"u8,
            _ => "hair"u8,
        });
        xml.WriteRaw("\">"u8);
        WriteColorElement(xml, "<color rgb=\""u8, edge.Color);
        xml.WriteRaw("</"u8);
        xml.WriteRaw(name);
        xml.WriteRaw(">"u8);
    }

    /// <summary>
    /// Writes a color element. The opening tag is passed in because the element is named after
    /// its role — <c>color</c>, <c>fgColor</c>, <c>bgColor</c> — while the attributes are the same.
    /// </summary>
    internal static void WriteColorElement(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> openRgb, ExcelColor color)
    {
        if (color.IsAuto)
        {
            WriteElementName(xml, openRgb);
            xml.WriteRaw(" auto=\"1\"/>"u8);
            return;
        }

        if (color.IsThemed)
        {
            WriteElementName(xml, openRgb);
            xml.WriteRaw(" theme=\""u8);
            xml.WriteInt32(color.ThemeIndex);
            if (color.Tint != 0)
            {
                xml.WriteRaw("\" tint=\""u8);
                xml.WriteDouble(color.Tint);
            }

            xml.WriteRaw("\"/>"u8);
            return;
        }

        xml.WriteRaw(openRgb);
        xml.WriteAttributeText(color.ToHex());
        xml.WriteRaw("\"/>"u8);
    }

    /// <summary>Writes just the element name out of an opening tag like <c>&lt;fgColor rgb="</c>.</summary>
    private static void WriteElementName(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> openTag)
    {
        int space = openTag.IndexOf((byte)' ');
        xml.WriteRaw(space < 0 ? openTag : openTag[..space]);
    }

    /// <summary>
    /// One <c>xf</c> record. A cell format points at its named style through
    /// <paramref name="styleXfId"/>; a master of <c>cellStyleXfs</c> passes null and carries
    /// no such attribute, being the thing pointed at.
    /// </summary>
    private static void WriteCellFormat(
        Utf8XmlWriter xml, CellStyle style, int numberFormatId, int fontId, int fillId, int borderId, int? styleXfId)
    {
        xml.WriteRaw("<xf numFmtId=\""u8);
        xml.WriteInt32(numberFormatId);
        xml.WriteRaw("\" fontId=\""u8);
        xml.WriteInt32(fontId);
        xml.WriteRaw("\" fillId=\""u8);
        xml.WriteInt32(fillId == 0 ? 0 : fillId + 1); // Skip the reserved gray125 entry.
        xml.WriteRaw("\" borderId=\""u8);
        xml.WriteInt32(borderId);
        if (styleXfId is { } xfId)
        {
            xml.WriteRaw("\" xfId=\""u8);
            xml.WriteInt32(xfId);
        }

        xml.WriteRaw("\""u8);
        if (numberFormatId != 0)
            xml.WriteRaw(" applyNumberFormat=\"1\""u8);
        if (fontId != 0)
            xml.WriteRaw(" applyFont=\"1\""u8);
        if (fillId != 0)
            xml.WriteRaw(" applyFill=\"1\""u8);
        if (borderId != 0)
            xml.WriteRaw(" applyBorder=\"1\""u8);

        bool hasAlignment = style.Alignment != Alignment.Default;
        bool hasProtection = style.Protection != CellProtection.Default;
        if (hasProtection)
            xml.WriteRaw(" applyProtection=\"1\""u8);
        if (!hasAlignment)
        {
            if (!hasProtection)
            {
                xml.WriteRaw("/>"u8);
                return;
            }

            xml.WriteRaw(">"u8);
            WriteProtection(xml, style.Protection);
            xml.WriteRaw("</xf>"u8);
            return;
        }

        Alignment alignment = style.Alignment;
        xml.WriteRaw(" applyAlignment=\"1\"><alignment"u8);
        if (alignment.Horizontal != HorizontalAlignment.General)
        {
            xml.WriteRaw(" horizontal=\""u8);
            xml.WriteRaw(alignment.Horizontal switch
            {
                HorizontalAlignment.Left => "left"u8,
                HorizontalAlignment.Center => "center"u8,
                HorizontalAlignment.Right => "right"u8,
                HorizontalAlignment.Justify => "justify"u8,
                HorizontalAlignment.Fill => "fill"u8,
                HorizontalAlignment.CenterContinuous => "centerContinuous"u8,
                _ => "distributed"u8,
            });
            xml.WriteRaw("\""u8);
        }

        if (alignment.Vertical != VerticalAlignment.Bottom)
        {
            xml.WriteRaw(" vertical=\""u8);
            xml.WriteRaw(alignment.Vertical switch
            {
                VerticalAlignment.Top => "top"u8,
                VerticalAlignment.Center => "center"u8,
                VerticalAlignment.Justify => "justify"u8,
                _ => "distributed"u8,
            });
            xml.WriteRaw("\""u8);
        }

        if (alignment.TextRotation != 0)
        {
            xml.WriteRaw(" textRotation=\""u8);
            xml.WriteInt32(alignment.TextRotation);
            xml.WriteRaw("\""u8);
        }

        if (alignment.WrapText)
            xml.WriteRaw(" wrapText=\"1\""u8);
        if (alignment.Indent > 0)
        {
            xml.WriteRaw(" indent=\""u8);
            xml.WriteInt32(alignment.Indent);
            xml.WriteRaw("\""u8);
        }

        if (alignment.ShrinkToFit)
            xml.WriteRaw(" shrinkToFit=\"1\""u8);
        if (alignment.ReadingOrder != ReadingOrder.Context)
        {
            xml.WriteRaw(" readingOrder=\""u8);
            xml.WriteInt32(alignment.ReadingOrder == ReadingOrder.LeftToRight ? 1 : 2);
            xml.WriteRaw("\""u8);
        }

        xml.WriteRaw("/>"u8);
        if (hasProtection)
            WriteProtection(xml, style.Protection);
        xml.WriteRaw("</xf>"u8);
    }

    /// <summary>
    /// Writes the <c>protection</c> child. Both attributes go out explicitly: leaving one off
    /// means its default, and the default of <c>locked</c> is the opposite of what a style that
    /// bothers to say anything about protection usually wants.
    /// </summary>
    private static void WriteProtection(Utf8XmlWriter xml, CellProtection protection)
    {
        xml.WriteRaw("<protection locked=\""u8);
        xml.WriteRaw(protection.Locked ? "1"u8 : "0"u8);
        xml.WriteRaw("\" hidden=\""u8);
        xml.WriteRaw(protection.Hidden ? "1"u8 : "0"u8);
        xml.WriteRaw("\"/>"u8);
    }
}
