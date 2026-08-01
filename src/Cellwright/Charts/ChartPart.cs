using System.Globalization;
using System.Xml;
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

    private static ChartView Read(XmlReader xml)
    {
        ChartKind kind = ChartKind.Unknown;
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
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Column;
                    break;

                case "barDir" when kind == ChartKind.Column:
                    if (xml.GetAttribute("val") == "bar")
                        kind = ChartKind.Bar;
                    break;

                case "lineChart" or "line3DChart" or "stockChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Line;
                    break;

                case "pieChart" or "pie3DChart" or "ofPieChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Pie;
                    break;

                case "doughnutChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Doughnut;
                    break;

                case "areaChart" or "area3DChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Area;
                    break;

                case "scatterChart" or "bubbleChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Scatter;
                    break;

                case "radarChart":
                    if (kind == ChartKind.Unknown)
                        kind = ChartKind.Radar;
                    break;

                case "ser":
                    series.Add(ReadSeries(xml));
                    advance = false;
                    break;

                default:
                    break;
            }
        }

        return new ChartView(kind, title, series);
    }

    /// <summary>Reads one series; the reader starts on <c>ser</c> and ends past its end tag.</summary>
    private static ChartSeries ReadSeries(XmlReader xml)
    {
        string? name = null;
        string? nameReference = null;
        string? categoriesReference = null;
        string? valuesReference = null;
        var categories = new List<string>();
        var values = new List<double?>();

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
                case "tx" or "cat" or "val" or "xVal" or "yVal":
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
            Name = name,
            NameReference = nameReference,
            CategoriesReference = categoriesReference,
            ValuesReference = valuesReference,
            Categories = categories,
            Values = values,
        };
    }
}
