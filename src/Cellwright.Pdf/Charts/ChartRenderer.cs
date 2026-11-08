using System.Globalization;
using Cellwright.Pdf.Layout;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Charts;

/// <summary>
/// Draws one chart into its rectangle on the page.
/// </summary>
internal sealed partial class ChartRenderer
{
    private const double Padding = 6;
    private const double LegendSwatch = 7;

    private static readonly PdfColor AxisColor = PdfColor.Rgb(0.35f, 0.35f, 0.35f);
    private static readonly PdfColor GridColor = PdfColor.Rgb(0.85f, 0.85f, 0.85f);

    private readonly PdfExportContext _context;
    private readonly TextAppearance _label;
    private readonly TextAppearance _title;

    internal ChartRenderer(PdfExportContext context, CellAppearances appearances)
    {
        _context = context;
        _label = appearances.Of(FontStyle.Default with { Size = 8 });
        _title = appearances.Of(FontStyle.Default with { Size = 12, Bold = true });
    }

    private PdfColor SeriesColor(int index) =>
        _context.ColorOf(ExcelColor.FromTheme(4 + (index % 6)), PdfColor.Rgb(0.3f, 0.5f, 0.7f));

    /// <summary>Draws a chart into a rectangle, appending to the page's items.</summary>
    public void Draw(ComposedPage page, Chart chart, double x, double y, double width, double height)
    {
        List<SeriesData> series = Materialize(chart);

        double top = DrawTitleIfPresent(page, chart.Title, x, y + Padding, width);
        double bottom = y + height - Padding;
        if (series.Count > 1 || chart.Kind is ChartKind.Pie or ChartKind.Doughnut)
            bottom -= DrawLegend(page, chart, series, x, bottom, width);

        var plot = new Rect(x + Padding, top, x + width - Padding, bottom);
        if (plot.Width <= 10 || plot.Height <= 10 || series.Count == 0)
            return;

        DrawPlotByKind(page, chart.Kind, series, plot);
    }

    private void DrawPlotByKind(ComposedPage page, ChartKind kind, List<SeriesData> series, Rect plot)
    {
        if (HasMixedCartesianKinds(kind, series))
        {
            DrawCombinedCartesian(page, series, plot);
            return;
        }

        switch (kind)
        {
            case ChartKind.Pie or ChartKind.Doughnut:
                DrawPie(page, series[0], plot, kind == ChartKind.Doughnut);
                break;
            case ChartKind.Bar:
                DrawBars(page, series, plot, horizontal: true);
                break;
            case ChartKind.Line:
                DrawLines(page, series, plot, filled: false);
                break;
            case ChartKind.Area:
                DrawLines(page, series, plot, filled: true);
                break;
            case ChartKind.Scatter:
                DrawScatter(page, series, plot);
                break;
            case ChartKind.Bubble:
                DrawBubble(page, series, plot);
                break;
            case ChartKind.Radar:
                DrawRadar(page, series, plot);
                break;
            default:
                DrawBars(page, series, plot, horizontal: false);
                break;
        }
    }

    private static bool HasMixedCartesianKinds(ChartKind chartKind, List<SeriesData> series) =>
        IsCartesianKind(chartKind) &&
        series.Any(s => s.Kind != chartKind) &&
        series.All(static s => IsCartesianKind(s.Kind));

    private static bool IsCartesianKind(ChartKind kind) =>
        kind is ChartKind.Column or ChartKind.Line or ChartKind.Area or ChartKind.Scatter or ChartKind.Bubble;

    private double DrawTitleIfPresent(ComposedPage page, string? title, double x, double top, double width)
    {
        if (title is not { Length: > 0 })
            return top;

        double measured = _title.Measure(title);
        page.Items.Add(new TextItem
        {
            X = x + ((width - measured) / 2),
            Y = top,
            Text = title,
            Appearance = _title,
            Width = measured,
            Height = _title.LineHeight,
        });
        return top + _title.LineHeight + Padding;
    }

    private sealed record SeriesData(
        int Index,
        ChartKind Kind,
        string? Name,
        IReadOnlyList<string> Categories,
        IReadOnlyList<double?> Values,
        IReadOnlyList<double?> BubbleSizes);

    private List<SeriesData> Materialize(Chart chart)
    {
        var series = new List<SeriesData>(chart.Series.Count);
        for (int index = 0; index < chart.Series.Count; index++)
        {
            ChartSeries source = chart.Series[index];
            IReadOnlyList<double?> values = source.Values;
            if (values.Count == 0 && source.ValuesReference is { } reference)
                values = ReadNumbers(reference);

            IReadOnlyList<string> categories = source.Categories;
            if (categories.Count == 0 && source.CategoriesReference is { } categoryReference)
                categories = ReadTexts(categoryReference);

            IReadOnlyList<double?> bubbleSizes = source.BubbleSizes;
            if (bubbleSizes.Count == 0 && source.BubbleSizesReference is { } bubbleReference)
                bubbleSizes = ReadNumbers(bubbleReference);

            ChartKind seriesKind = source.Kind != ChartKind.Unknown ? source.Kind : chart.Kind;
            if (values.Count > 0)
                series.Add(new SeriesData(index, seriesKind, source.Name, categories, values, bubbleSizes));
        }

        return series;
    }

    private IReadOnlyList<double?> ReadNumbers(string reference)
    {
        var values = new List<double?>();
        foreach (CellValue cell in ReadCells(reference))
            values.Add(cell.TryGetNumber(out double number) ? number : null);
        return values;
    }

    private IReadOnlyList<string> ReadTexts(string reference)
    {
        var texts = new List<string>();
        foreach (CellValue cell in ReadCells(reference))
            texts.Add(cell.ToString());
        return texts;
    }

    private List<CellValue> ReadCells(string reference)
    {
        var cells = new List<CellValue>();
        if (SheetRangeResolver.Resolve(_context.Source, reference) is not ({ } sheet, { } range))
            return cells;

        for (int row = range.Start.Row; row <= range.End.Row; row++)
        {
            for (int column = range.Start.Column; column <= range.End.Column; column++)
                cells.Add(sheet.Values.Get(row, column));
        }

        return cells;
    }

    private readonly record struct Rect(double Left, double Top, double Right, double Bottom)
    {
        public double Width => Right - Left;

        public double Height => Bottom - Top;
    }

    private double DrawLegend(ComposedPage page, Chart chart, List<SeriesData> series, double x, double bottom, double width)
    {
        List<(string Text, PdfColor Color)> entries = BuildLegendEntries(chart.Kind, series);
        double rowHeight = Math.Max(_label.LineHeight, LegendSwatch) + 2;
        double cursor = x + Padding;
        double top = bottom - rowHeight;

        foreach ((string text, PdfColor color) in entries)
        {
            double entryWidth = LegendSwatch + 3 + _label.Measure(text) + 10;
            if (cursor + entryWidth > x + width - Padding)
                break;

            page.Items.Add(new FillItem
            {
                X = cursor,
                Y = top + ((rowHeight - LegendSwatch) / 2),
                Width = LegendSwatch,
                Height = LegendSwatch,
                Color = color,
            });

            page.Items.Add(new TextItem
            {
                X = cursor + LegendSwatch + 3,
                Y = top + ((rowHeight - _label.LineHeight) / 2),
                Text = text,
                Appearance = _label,
                Width = _label.Measure(text),
                Height = _label.LineHeight,
            });

            cursor += entryWidth;
        }

        return rowHeight + 2;
    }

    private List<(string Text, PdfColor Color)> BuildLegendEntries(ChartKind kind, List<SeriesData> series)
    {
        List<(string Text, PdfColor Color)> entries = [];
        if (kind is ChartKind.Pie or ChartKind.Doughnut && series.Count > 0)
        {
            for (int i = 0; i < series[0].Values.Count; i++)
            {
                string label = i < series[0].Categories.Count ? series[0].Categories[i] : $"{i + 1}";
                entries.Add((label, SeriesColor(i)));
            }
        }
        else
        {
            for (int i = 0; i < series.Count; i++)
                entries.Add((series[i].Name ?? $"Series {i + 1}", SeriesColor(i)));
        }

        return entries;
    }

    private static (double Low, double High) Extent(List<SeriesData> series)
    {
        double low = 0;
        double high = 0;
        foreach (SeriesData data in series)
        {
            foreach (double? value in data.Values)
            {
                if (value is not { } number)
                    continue;

                low = Math.Min(low, number);
                high = Math.Max(high, number);
            }
        }

        if (high == low)
            high = low + 1;
        return (low, high);
    }

    private static double NiceStep(double range)
    {
        double rough = range / 4;
        double magnitude = Math.Pow(10, Math.Floor(Math.Log10(Math.Max(rough, 1e-9))));
        double fraction = rough / magnitude;
        return (fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10) * magnitude;
    }

    private static string Format(double value) =>
        Math.Abs(value) >= 1000
            ? value.ToString("#,##0", CultureInfo.InvariantCulture)
            : value.ToString("0.##", CultureInfo.InvariantCulture);
}
