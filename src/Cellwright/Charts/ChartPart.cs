using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Cellwright.Xml;

namespace Cellwright;

/// <summary>
/// Parses a <c>chartSpace</c> part into the view the model shows (§21.2).
/// </summary>
/// <remarks>
/// Deliberately a view, not a round-trip: what is parsed here is what <see cref="Chart"/>
/// exposes, and everything else — axes, styling, layout, the second and third plot groups —
/// stays in the raw markup, which is what a save writes back. Parsing by local name under the
/// chart namespace keeps files from producers with unusual prefixes readable, the same rule the
/// sheet-rename retargeting follows.
/// </remarks>
internal static class ChartPart
{
    /// <summary>Both spellings of the chart namespace: transitional and strict.</summary>
    private static bool IsChartNamespace(string ns) =>
        ns is "http://schemas.openxmlformats.org/drawingml/2006/chart"
            or "http://purl.oclc.org/ooxml/drawingml/chart";

    /// <summary>Both spellings of the DrawingML main namespace, where the title's text lives.</summary>
    private static bool IsDrawingNamespace(string ns) =>
        ns is "http://schemas.openxmlformats.org/drawingml/2006/main"
            or "http://purl.oclc.org/ooxml/drawingml/main";

    /// <summary>Parses the part; a part that will not parse is an empty view, not an error.</summary>
    public static ChartView Parse(ReadOnlyMemory<byte> content)
    {
        if (content.IsEmpty)
            return ChartView.Empty;

        try
        {
            using var stream = new MemoryStream(content.ToArray(), writable: false);
            using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
            return Read(xml);
        }
        catch (XmlException)
        {
            return ChartView.Empty;
        }
    }

    /// <summary>
    /// Rewrites chart formula elements that name a worksheet. The full chartSpace remains the
    /// source of truth; only <c>c:f</c> text changes, and external-book references are left alone
    /// by <see cref="FormulaText.RenameSheet(string?, string, string)"/>.
    /// </summary>
    /// <param name="content">Raw chartSpace bytes.</param>
    /// <param name="oldName">Source worksheet name.</param>
    /// <param name="newName">Copied worksheet name.</param>
    /// <param name="parsed">Whether the chartSpace was well-formed XML.</param>
    /// <param name="changed">Whether any formula text was rewritten.</param>
    /// <returns>The original memory when unchanged, otherwise rewritten UTF-8 XML.</returns>
    internal static ReadOnlyMemory<byte> RetargetSheetName(
        ReadOnlyMemory<byte> content,
        string oldName,
        string newName,
        out bool parsed,
        out bool changed)
    {
        parsed = true;
        changed = false;
        if (content.IsEmpty)
            return content;

        XDocument document;
        try
        {
            using var input = new MemoryStream(content.ToArray(), writable: false);
            using var xml = XmlReader.Create(input, XmlDefaults.ReaderSettings);
            document = XDocument.Load(xml, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            parsed = false;
            return content;
        }

        foreach (XElement formula in document.Descendants().Where(static element =>
                     element.Name.LocalName == "f" && IsChartNamespace(element.Name.NamespaceName)))
        {
            string rewritten = FormulaText.RenameSheet(formula.Value, oldName, newName)!;
            if (string.Equals(formula.Value, rewritten, StringComparison.Ordinal))
                continue;

            formula.Value = rewritten;
            changed = true;
        }

        if (!changed)
            return content;

        using var output = new MemoryStream(content.Length);
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration is null,
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings))
            document.Save(writer);
        return output.ToArray();
    }

    /// <summary>
    /// Rewrites chart formula elements (<c>c:f</c>) when rows or columns are inserted or deleted.
    /// </summary>
    internal static ReadOnlyMemory<byte> RetargetShift(
        ReadOnlyMemory<byte> content,
        in ReferenceShift shift,
        string ownSheet,
        out bool changed)
    {
        changed = false;
        if (content.IsEmpty)
            return content;

        XDocument document;
        try
        {
            using var input = new MemoryStream(content.ToArray(), writable: false);
            using var xml = XmlReader.Create(input, XmlDefaults.ReaderSettings);
            document = XDocument.Load(xml, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return content;
        }

        foreach (XElement formula in document.Descendants().Where(static element =>
                     element.Name.LocalName == "f" && IsChartNamespace(element.Name.NamespaceName)))
        {
            string rewritten = FormulaText.Shift(formula.Value, shift, ownSheet)!;
            if (string.Equals(formula.Value, rewritten, StringComparison.Ordinal))
                continue;

            formula.Value = rewritten;
            changed = true;
        }

        if (!changed)
            return content;

        using var output = new MemoryStream(content.Length);
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration is null,
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings))
            document.Save(writer);
        return output.ToArray();
    }

    private static ChartView Read(XmlReader xml)
    {
        ChartKind kind = ChartKind.Unknown;
        ChartKind groupKind = ChartKind.Unknown;
        int groupSeriesStart = 0;
        string? title = null;
        var series = new List<ChartSeries>();
        bool inTitle = false;
        int titleDepth = 0;

        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && inTitle && xml.Depth <= titleDepth)
            {
                inTitle = false;
                continue;
            }

            if (xml.NodeType != XmlNodeType.Element)
                continue;

            // The title's written-in text arrives as DrawingML runs — a namespace of their
            // own — and their concatenation is the title. A formula-driven title has none.
            if (inTitle && xml.LocalName == "t" && IsDrawingNamespace(xml.NamespaceURI))
            {
                title = (title ?? "") + xml.ReadElementContentAsString();
                advance = false;
                continue;
            }

            if (!IsChartNamespace(xml.NamespaceURI))
                continue;

            switch (xml.LocalName)
            {
                case "title":
                    inTitle = true;
                    titleDepth = xml.Depth;
                    break;

                case "barChart" or "bar3DChart":
                    // Direction decides bar against column; the child comes right after.
                    groupKind = ChartKind.Column;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Column;
                    break;

                case "barDir" when groupKind == ChartKind.Column:
                    if (xml.GetAttribute("val") == "bar")
                    {
                        groupKind = ChartKind.Bar;
                        if (kind == ChartKind.Column && groupSeriesStart == 0)
                            kind = ChartKind.Bar;
                        for (int i = groupSeriesStart; i < series.Count; i++)
                            series[i] = series[i] with { Kind = ChartKind.Bar };
                    }
                    break;

                case "lineChart" or "line3DChart" or "stockChart":
                    groupKind = ChartKind.Line;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Line;
                    break;

                case "pieChart" or "pie3DChart" or "ofPieChart":
                    groupKind = ChartKind.Pie;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Pie;
                    break;

                case "doughnutChart":
                    groupKind = ChartKind.Doughnut;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Doughnut;
                    break;

                case "areaChart" or "area3DChart":
                    groupKind = ChartKind.Area;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Area;
                    break;

                case "scatterChart":
                    groupKind = ChartKind.Scatter;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Scatter;
                    break;

                case "bubbleChart":
                    groupKind = ChartKind.Bubble;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Bubble;
                    break;

                case "radarChart":
                    groupKind = ChartKind.Radar;
                    groupSeriesStart = series.Count;
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Radar;
                    break;

                case "ser":
                    series.Add(ReadSeries(xml, groupKind));
                    advance = false;
                    break;

                default:
                    break;
            }
        }

        return new ChartView(kind, title, series);
    }

    /// <summary>Reads one series; the reader starts on <c>ser</c> and ends past its end tag.</summary>
    private static ChartSeries ReadSeries(XmlReader xml, ChartKind groupKind)
    {
        string? name = null;
        string? nameReference = null;
        string? categoriesReference = null;
        string? valuesReference = null;
        string? bubbleSizesReference = null;
        var categories = new List<string>();
        var values = new List<double?>();
        var bubbleSizes = new List<double?>();

        // Which block the reader is inside decides what an f or a pt belongs to.
        string block = "";
        int depth = xml.Depth;

        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.Depth <= depth && xml.LocalName == "ser")
                break;

            if (xml.NodeType != XmlNodeType.Element || !IsChartNamespace(xml.NamespaceURI))
                continue;

            switch (xml.LocalName)
            {
                case "tx" or "cat" or "val" or "xVal" or "yVal" or "bubbleSize":
                    block = xml.LocalName;
                    break;

                case "f":
                    {
                        string reference = xml.ReadElementContentAsString();
                        advance = false;
                        switch (block)
                        {
                            case "tx":
                                nameReference = reference;
                                break;
                            case "cat" or "xVal":
                                categoriesReference = reference;
                                break;
                            case "val" or "yVal":
                                valuesReference = reference;
                                break;
                            case "bubbleSize":
                                bubbleSizesReference = reference;
                                break;
                            default:
                                break;
                        }

                        break;
                    }

                case "v":
                    {
                        string text = xml.ReadElementContentAsString();
                        advance = false;
                        switch (block)
                        {
                            case "tx":
                                name = text;
                                break;
                            case "cat" or "xVal":
                                categories.Add(text);
                                break;
                            case "val" or "yVal":
                                values.Add(double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                                    ? number
                                    : null);
                                break;
                            case "bubbleSize":
                                bubbleSizes.Add(double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double size)
                                    ? size
                                    : null);
                                break;
                            default:
                                break;
                        }

                        break;
                    }

                default:
                    break;
            }
        }

        return new ChartSeries
        {
            Kind = groupKind,
            Name = name,
            NameReference = nameReference,
            CategoriesReference = categoriesReference,
            ValuesReference = valuesReference,
            BubbleSizesReference = bubbleSizesReference,
            Categories = categories,
            Values = values,
            BubbleSizes = bubbleSizes,
        };
    }
}
