using Cellwright.Xml;

namespace Cellwright;

/// <summary>One series of a chart being created: where its data comes from.</summary>
/// <param name="ValuesReference">Reference the values come from, e.g. <c>Данные!$B$2:$B$10</c>.</param>
/// <param name="CategoriesReference">Reference the category labels come from, or <see langword="null"/>.</param>
/// <param name="Name">The series name, or <see langword="null"/> to leave it to the reader.</param>
public sealed record ChartSeriesDefinition(
    string ValuesReference,
    string? CategoriesReference = null,
    string? Name = null)
{
    /// <summary>Reference the bubble sizes come from, or <see langword="null"/>.</summary>
    public string? BubbleSizesReference { get; init; }

    /// <summary>Per-series kind for a combined chart, or <see langword="null"/> for the chart's kind.</summary>
    public ChartKind? Kind { get; init; }

    /// <summary>Creates a series with optional bubble sizes and a per-series plot kind.</summary>
    public ChartSeriesDefinition(
        string ValuesReference, string? CategoriesReference, string? Name,
        string? BubbleSizesReference, ChartKind? Kind = null)
        : this(ValuesReference, CategoriesReference, Name)
    {
        this.BubbleSizesReference = BubbleSizesReference;
        this.Kind = Kind;
    }

    /// <summary>Creates a series with its own plot kind for a combined chart.</summary>
    public ChartSeriesDefinition(
        string ValuesReference, string? CategoriesReference, string? Name, ChartKind Kind)
        : this(ValuesReference, CategoriesReference, Name, null, Kind) { }

    /// <summary>Deconstructs the series including its bubble sizes and plot kind.</summary>
    public void Deconstruct(
        out string ValuesReference, out string? CategoriesReference, out string? Name,
        out string? BubbleSizesReference, out ChartKind? Kind)
    {
        ValuesReference = this.ValuesReference;
        CategoriesReference = this.CategoriesReference;
        Name = this.Name;
        BubbleSizesReference = this.BubbleSizesReference;
        Kind = this.Kind;
    }
}

/// <summary>
/// Generates the <c>chartSpace</c> part of a chart the model creates.
/// </summary>
/// <remarks>
/// The smallest markup Excel accepts and draws: the plot group for the kind, the series with
/// their references, and — for the kinds that have them — a category and a value axis wired to
/// each other. The caches beside each reference are filled from the live cells when the
/// reference resolves, because a chart without caches shows empty until its first
/// recalculation in Excel, and other readers never recalculate at all.
/// </remarks>
internal static class ChartWriter
{
    /// <summary>Fixed axis identifiers; they only have to agree with each other.</summary>
    private const string CategoryAxisId = "111111111";
    private const string ValueAxisId = "222222222";

    /// <summary>The kinds a chart can be created as.</summary>
    public static bool CanCreate(ChartKind kind) => kind is
        ChartKind.Column or ChartKind.Bar or ChartKind.Line or
        ChartKind.Pie or ChartKind.Doughnut or ChartKind.Area or
        ChartKind.Scatter or ChartKind.Bubble or ChartKind.Radar;

    /// <summary>Builds the part for a new chart.</summary>
    /// <param name="workbook">The workbook the series references speak of, for the caches.</param>
    /// <param name="kind">What to plot the series as.</param>
    /// <param name="series">The series, in plot order.</param>
    /// <param name="title">The chart title, or <see langword="null"/> for none.</param>
    public static byte[] Build(
        Workbook workbook, ChartKind kind, IReadOnlyList<ChartSeriesDefinition> series, string? title)
    {
        string markup = Utf8XmlWriter.Render(xml =>
        {
            xml.WriteDeclaration();
            xml.WriteRaw("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""u8);
            xml.WriteRaw(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""u8);
            xml.WriteRaw(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8);
            xml.WriteRaw("<c:chart>"u8);

            if (title is not null)
            {
                xml.WriteRaw("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>"u8);
                xml.WriteText(title);
                xml.WriteRaw("</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>"u8);
                xml.WriteRaw("<c:autoTitleDeleted val=\"0\"/>"u8);
            }

            xml.WriteRaw("<c:plotArea><c:layout/>"u8);
            bool anyAxes = false;
            int index = 0;
            while (index < series.Count)
            {
                ChartKind groupKind = series[index].Kind is { } overrideKind && CanCreate(overrideKind)
                    ? overrideKind
                    : kind;
                int groupEnd = index + 1;
                while (groupEnd < series.Count)
                {
                    ChartKind nextKind = series[groupEnd].Kind is { } nextOverride && CanCreate(nextOverride)
                        ? nextOverride
                        : kind;
                    if (nextKind != groupKind)
                        break;
                    groupEnd++;
                }

                WritePlotGroup(xml, workbook, groupKind, series, index, groupEnd);
                anyAxes |= HasAxes(groupKind);
                index = groupEnd;
            }

            if (anyAxes)
                WriteAxes(xml);
            xml.WriteRaw("</c:plotArea>"u8);

            xml.WriteRaw("<c:plotVisOnly val=\"1\"/>"u8);
            xml.WriteRaw("</c:chart></c:chartSpace>"u8);
        });

        return System.Text.Encoding.UTF8.GetBytes(markup);
    }

    /// <summary>Whether the kind plots against a pair of axes; a pie has none.</summary>
    private static bool HasAxes(ChartKind kind) => kind is not (ChartKind.Pie or ChartKind.Doughnut);

    /// <summary>The category and value axes, each naming the other as what it crosses.</summary>
    private static void WriteAxes(Utf8XmlWriter xml)
    {
        xml.WriteRaw("<c:catAx><c:axId val=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(CategoryAxisId));
        xml.WriteRaw("\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>"u8);
        xml.WriteRaw("<c:axPos val=\"b\"/><c:crossAx val=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(ValueAxisId));
        xml.WriteRaw("\"/></c:catAx>"u8);

        xml.WriteRaw("<c:valAx><c:axId val=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(ValueAxisId));
        xml.WriteRaw("\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>"u8);
        xml.WriteRaw("<c:axPos val=\"l\"/><c:crossAx val=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(CategoryAxisId));
        xml.WriteRaw("\"/></c:valAx>"u8);
    }

    private static void WritePlotGroup(
        Utf8XmlWriter xml,
        Workbook workbook,
        ChartKind kind,
        IReadOnlyList<ChartSeriesDefinition> series,
        int startIndex,
        int endIndex)
    {
        ReadOnlySpan<byte> element = kind switch
        {
            ChartKind.Column or ChartKind.Bar => "c:barChart"u8,
            ChartKind.Line => "c:lineChart"u8,
            ChartKind.Pie => "c:pieChart"u8,
            ChartKind.Doughnut => "c:doughnutChart"u8,
            ChartKind.Scatter => "c:scatterChart"u8,
            ChartKind.Bubble => "c:bubbleChart"u8,
            ChartKind.Radar => "c:radarChart"u8,
            _ => "c:areaChart"u8,
        };

        xml.WriteRaw("<"u8);
        xml.WriteRaw(element);
        xml.WriteRaw(">"u8);

        switch (kind)
        {
            case ChartKind.Column:
                xml.WriteRaw("<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>"u8);
                break;
            case ChartKind.Bar:
                xml.WriteRaw("<c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>"u8);
                break;
            case ChartKind.Line or ChartKind.Area:
                xml.WriteRaw("<c:grouping val=\"standard\"/>"u8);
                break;
            case ChartKind.Pie or ChartKind.Doughnut:
                xml.WriteRaw("<c:varyColors val=\"1\"/>"u8);
                break;
            case ChartKind.Scatter:
                xml.WriteRaw("<c:scatterStyle val=\"lineMarker\"/>"u8);
                break;
            case ChartKind.Bubble:
                xml.WriteRaw("<c:varyColors val=\"0\"/>"u8);
                break;
            case ChartKind.Radar:
                xml.WriteRaw("<c:radarStyle val=\"standard\"/>"u8);
                break;
            default:
                break;
        }

        for (int i = startIndex; i < endIndex; i++)
            WriteSeries(xml, workbook, kind, series[i], i);

        if (kind == ChartKind.Doughnut)
            xml.WriteRaw("<c:holeSize val=\"50\"/>"u8);

        if (HasAxes(kind))
        {
            xml.WriteRaw("<c:axId val=\""u8);
            xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(CategoryAxisId));
            xml.WriteRaw("\"/><c:axId val=\""u8);
            xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(ValueAxisId));
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</"u8);
        xml.WriteRaw(element);
        xml.WriteRaw(">"u8);
    }

    private static void WriteSeries(
        Utf8XmlWriter xml, Workbook workbook, ChartKind kind, ChartSeriesDefinition series, int index)
    {
        bool isXY = kind is ChartKind.Scatter or ChartKind.Bubble;
        xml.WriteRaw("<c:ser><c:idx val=\""u8);
        xml.WriteInt32(index);
        xml.WriteRaw("\"/><c:order val=\""u8);
        xml.WriteInt32(index);
        xml.WriteRaw("\"/>"u8);

        if (series.Name is { Length: > 0 } name)
        {
            xml.WriteRaw("<c:tx><c:v>"u8);
            xml.WriteText(name);
            xml.WriteRaw("</c:v></c:tx>"u8);
        }

        if (series.CategoriesReference is { Length: > 0 } categories)
        {
            xml.WriteRaw(isXY ? "<c:xVal><c:strRef><c:f>"u8 : "<c:cat><c:strRef><c:f>"u8);
            xml.WriteText(categories);
            xml.WriteRaw("</c:f>"u8);
            WriteStringCache(xml, workbook, categories);
            xml.WriteRaw(isXY ? "</c:strRef></c:xVal>"u8 : "</c:strRef></c:cat>"u8);
        }

        xml.WriteRaw(isXY ? "<c:yVal><c:numRef><c:f>"u8 : "<c:val><c:numRef><c:f>"u8);
        xml.WriteText(series.ValuesReference);
        xml.WriteRaw("</c:f>"u8);
        WriteNumberCache(xml, workbook, series.ValuesReference);
        xml.WriteRaw(isXY ? "</c:numRef></c:yVal>"u8 : "</c:numRef></c:val>"u8);

        if (kind == ChartKind.Bubble)
        {
            string bubbleRef = series.BubbleSizesReference is { Length: > 0 } explicitRef
                ? explicitRef
                : series.ValuesReference;
            xml.WriteRaw("<c:bubbleSize><c:numRef><c:f>"u8);
            xml.WriteText(bubbleRef);
            xml.WriteRaw("</c:f>"u8);
            WriteNumberCache(xml, workbook, bubbleRef);
            xml.WriteRaw("</c:numRef></c:bubbleSize>"u8);
        }

        xml.WriteRaw("</c:ser>"u8);
    }

    /// <summary>Writes the cached values of a reference, read from the live cells.</summary>
    private static void WriteNumberCache(Utf8XmlWriter xml, Workbook workbook, string reference)
    {
        List<CellValue> cells = Read(workbook, reference);
        xml.WriteRaw("<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\""u8);
        xml.WriteInt32(cells.Count);
        xml.WriteRaw("\"/>"u8);

        for (int i = 0; i < cells.Count; i++)
        {
            if (!cells[i].TryGetNumber(out double number))
                continue;

            xml.WriteRaw("<c:pt idx=\""u8);
            xml.WriteInt32(i);
            xml.WriteRaw("\"><c:v>"u8);
            xml.WriteDouble(number);
            xml.WriteRaw("</c:v></c:pt>"u8);
        }

        xml.WriteRaw("</c:numCache>"u8);
    }

    private static void WriteStringCache(Utf8XmlWriter xml, Workbook workbook, string reference)
    {
        List<CellValue> cells = Read(workbook, reference);
        xml.WriteRaw("<c:strCache><c:ptCount val=\""u8);
        xml.WriteInt32(cells.Count);
        xml.WriteRaw("\"/>"u8);

        for (int i = 0; i < cells.Count; i++)
        {
            if (cells[i].IsBlank)
                continue;

            xml.WriteRaw("<c:pt idx=\""u8);
            xml.WriteInt32(i);
            xml.WriteRaw("\"><c:v>"u8);
            xml.WriteText(cells[i].ToString());
            xml.WriteRaw("</c:v></c:pt>"u8);
        }

        xml.WriteRaw("</c:strCache>"u8);
    }

    /// <summary>
    /// The cells a reference names, in row-major order, or nothing when it does not resolve —
    /// the cache is then empty and the reader computes for itself.
    /// </summary>
    private static List<CellValue> Read(Workbook workbook, string reference)
    {
        var cells = new List<CellValue>();
        if (SheetRangeResolver.Resolve(workbook, reference) is not ({ } sheet, { } range))
            return cells;

        // A whole-column reference would be a million points; clamp to the used range the way
        // a chart clamps its own plotting.
        int lastRow = Math.Min(range.End.Row, sheet.UsedRange?.End.Row ?? range.Start.Row);
        int lastColumn = Math.Min(range.End.Column, sheet.UsedRange?.End.Column ?? range.Start.Column);

        for (int row = range.Start.Row; row <= lastRow; row++)
        {
            for (int column = range.Start.Column; column <= lastColumn; column++)
                cells.Add(sheet.Values.Get(row, column));
        }

        return cells;
    }
}
