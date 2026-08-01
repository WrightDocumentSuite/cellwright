using Cellwright.Pdf.Layout;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Charts;

/// <summary>
/// Draws one chart into its rectangle on the page.
/// </summary>
/// <remarks>
/// Drawn from the model rather than from the markup: the series hand over their points, and the
/// renderer lays out a title, a plot area, axes where the kind has them, and a legend when there
/// is more than one series to tell apart. The data comes from the caches the file carries, and
/// from the live cells when a series arrived without them — a chart from Excel always caches,
/// but a chart from another producer may not.
/// </remarks>
internal sealed class ChartRenderer
{
    private const double Padding = 6;
    private const double TickLength = 3;
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

    /// <summary>The colour of a series: the theme's accents, round and round.</summary>
    private PdfColor SeriesColor(int index) =>
        _context.ColorOf(ExcelColor.FromTheme(4 + (index % 6)), PdfColor.Rgb(0.3f, 0.5f, 0.7f));

    /// <summary>Draws a chart into a rectangle, appending to the page's items.</summary>
    /// <param name="page">The page being composed.</param>
    /// <param name="chart">The chart to draw.</param>
    /// <param name="x">Left edge of its frame, in the page's own space.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Frame width.</param>
    /// <param name="height">Frame height.</param>
    public void Draw(ComposedPage page, Chart chart, double x, double y, double width, double height)
    {
        List<SeriesData> series = Materialize(chart);

        double top = y + Padding;
        if (chart.Title is { Length: > 0 } title)
        {
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
            top += _title.LineHeight + Padding;
        }

        double bottom = y + height - Padding;
        if (series.Count > 1 || chart.Kind is ChartKind.Pie or ChartKind.Doughnut)
            bottom -= DrawLegend(page, chart, series, x, y + height - Padding, width);

        var plot = new Rect(x + Padding, top, x + width - Padding, bottom);
        if (plot.Width <= 10 || plot.Height <= 10 || series.Count == 0)
            return;

        switch (chart.Kind)
        {
            case ChartKind.Pie or ChartKind.Doughnut:
                DrawPie(page, series[0], plot, chart.Kind == ChartKind.Doughnut);
                break;

            case ChartKind.Bar:
                DrawBars(page, series, plot, horizontal: true);
                break;

            case ChartKind.Line or ChartKind.Scatter:
                DrawLines(page, series, plot, filled: false);
                break;

            case ChartKind.Area:
                DrawLines(page, series, plot, filled: true);
                break;

            default:
                DrawBars(page, series, plot, horizontal: false);
                break;
        }
    }

    /// <summary>A series with its points in hand, whatever they came from.</summary>
    private sealed record SeriesData(string? Name, IReadOnlyList<string> Categories, IReadOnlyList<double?> Values);

    /// <summary>The chart's series with their points: the caches, or the live cells without them.</summary>
    private List<SeriesData> Materialize(Chart chart)
    {
        var series = new List<SeriesData>(chart.Series.Count);
        foreach (ChartSeries source in chart.Series)
        {
            IReadOnlyList<double?> values = source.Values;
            if (values.Count == 0 && source.ValuesReference is { } reference)
                values = ReadNumbers(reference);

            IReadOnlyList<string> categories = source.Categories;
            if (categories.Count == 0 && source.CategoriesReference is { } categoryReference)
                categories = ReadTexts(categoryReference);

            if (values.Count > 0)
                series.Add(new SeriesData(source.Name, categories, values));
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

    /// <summary>Draws the legend along the bottom edge; returns the height it took.</summary>
    private double DrawLegend(ComposedPage page, Chart chart, List<SeriesData> series, double x, double bottom, double width)
    {
        // A pie's legend names its slices; anything else names its series.
        List<(string Text, PdfColor Color)> entries = [];
        if (chart.Kind is ChartKind.Pie or ChartKind.Doughnut && series.Count > 0)
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

    /// <summary>The lowest and highest a value axis has to reach, zero always included.</summary>
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

    /// <summary>A tick step that lands on round numbers: 1, 2 or 5 times a power of ten.</summary>
    private static double NiceStep(double range)
    {
        double rough = range / 4;
        double magnitude = Math.Pow(10, Math.Floor(Math.Log10(Math.Max(rough, 1e-9))));
        double fraction = rough / magnitude;
        return (fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10) * magnitude;
    }

    /// <summary>Draws the value axis: its line, its round-numbered ticks, and their grid lines.</summary>
    /// <returns>How to map a value onto the axis, after room for the labels was set aside.</returns>
    private (Rect Plot, Func<double, double> Map) ValueAxis(
        ComposedPage page, List<SeriesData> series, Rect area, bool horizontal)
    {
        (double low, double high) = Extent(series);
        double step = NiceStep(high - low);
        low = Math.Floor(low / step) * step;
        high = Math.Ceiling(high / step) * step;

        // Room for the widest tick label, on the left for a vertical axis and below for a
        // horizontal one; the category labels take the other side.
        double labelRoom = 0;
        for (double tick = low; tick <= high + (step / 2); tick += step)
            labelRoom = Math.Max(labelRoom, _label.Measure(Format(tick)));

        Rect plot = horizontal
            ? new Rect(area.Left, area.Top, area.Right, area.Bottom - _label.LineHeight - 2 - _label.LineHeight)
            : new Rect(area.Left + labelRoom + 4, area.Top, area.Right, area.Bottom - _label.LineHeight - 2);

        double span = high - low;
        Func<double, double> map = horizontal
            ? value => plot.Left + ((value - low) / span * plot.Width)
            : value => plot.Bottom - ((value - low) / span * plot.Height);

        for (double tick = low; tick <= high + (step / 2); tick += step)
        {
            string label = Format(tick);
            double at = map(tick);

            if (horizontal)
            {
                page.Items.Add(new StrokeItem
                {
                    X = at, Y = plot.Top, X2 = at, Y2 = plot.Bottom, Thickness = 0.4, Color = GridColor,
                });
                page.Items.Add(new TextItem
                {
                    X = at - (_label.Measure(label) / 2),
                    Y = plot.Bottom + 2,
                    Text = label,
                    Appearance = _label,
                    Width = _label.Measure(label),
                    Height = _label.LineHeight,
                });
            }
            else
            {
                page.Items.Add(new StrokeItem
                {
                    X = plot.Left, Y = at, X2 = plot.Right, Y2 = at, Thickness = 0.4, Color = GridColor,
                });
                page.Items.Add(new TextItem
                {
                    X = plot.Left - 4 - _label.Measure(label),
                    Y = at - (_label.LineHeight / 2),
                    Text = label,
                    Appearance = _label,
                    Width = _label.Measure(label),
                    Height = _label.LineHeight,
                });
            }
        }

        // The axis lines themselves, over the grid.
        page.Items.Add(new StrokeItem
        {
            X = plot.Left, Y = plot.Bottom, X2 = plot.Right, Y2 = plot.Bottom, Thickness = 0.8, Color = AxisColor,
        });
        page.Items.Add(new StrokeItem
        {
            X = plot.Left, Y = plot.Top, X2 = plot.Left, Y2 = plot.Bottom, Thickness = 0.8, Color = AxisColor,
        });

        return (plot, map);
    }

    private static string Format(double value) =>
        Math.Abs(value) >= 1000
            ? value.ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture)
            : value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>Category labels along the bottom, one under each band that has the room.</summary>
    private void DrawCategoryLabels(
        ComposedPage page, IReadOnlyList<string> categories, Rect plot, int count, double under)
    {
        if (count == 0)
            return;

        double band = plot.Width / count;
        for (int i = 0; i < count && i < categories.Count; i++)
        {
            string label = categories[i];
            double measured = _label.Measure(label);
            if (measured > band)
                continue;

            page.Items.Add(new TextItem
            {
                X = plot.Left + (band * i) + ((band - measured) / 2),
                Y = under,
                Text = label,
                Appearance = _label,
                Width = measured,
                Height = _label.LineHeight,
            });
        }
    }

    private void DrawBars(ComposedPage page, List<SeriesData> series, Rect area, bool horizontal)
    {
        (Rect plot, Func<double, double> map) = ValueAxis(page, series, area, horizontal);

        int points = 0;
        foreach (SeriesData data in series)
            points = Math.Max(points, data.Values.Count);
        if (points == 0)
            return;

        double baseline = map(0);

        if (horizontal)
        {
            double band = plot.Height / points;
            double bar = band * 0.7 / series.Count;
            for (int s = 0; s < series.Count; s++)
            {
                PdfColor color = SeriesColor(s);
                for (int i = 0; i < series[s].Values.Count; i++)
                {
                    if (series[s].Values[i] is not { } value)
                        continue;

                    double at = map(value);
                    page.Items.Add(new FillItem
                    {
                        X = Math.Min(baseline, at),
                        Y = plot.Top + (band * i) + (band * 0.15) + (bar * s),
                        Width = Math.Abs(at - baseline),
                        Height = bar,
                        Color = color,
                    });
                }
            }

            DrawCategoryLabels(
                page, series[0].Categories, plot with { Left = plot.Left, Right = plot.Right }, points,
                plot.Bottom + 2 + _label.LineHeight);
        }
        else
        {
            double band = plot.Width / points;
            double bar = band * 0.7 / series.Count;
            for (int s = 0; s < series.Count; s++)
            {
                PdfColor color = SeriesColor(s);
                for (int i = 0; i < series[s].Values.Count; i++)
                {
                    if (series[s].Values[i] is not { } value)
                        continue;

                    double at = map(value);
                    page.Items.Add(new FillItem
                    {
                        X = plot.Left + (band * i) + (band * 0.15) + (bar * s),
                        Y = Math.Min(baseline, at),
                        Width = bar,
                        Height = Math.Abs(at - baseline),
                        Color = color,
                    });
                }
            }

            DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
        }
    }

    private void DrawLines(ComposedPage page, List<SeriesData> series, Rect area, bool filled)
    {
        (Rect plot, Func<double, double> map) = ValueAxis(page, series, area, horizontal: false);

        int points = 0;
        foreach (SeriesData data in series)
            points = Math.Max(points, data.Values.Count);
        if (points == 0)
            return;

        double band = plot.Width / points;

        for (int s = 0; s < series.Count; s++)
        {
            PdfColor color = SeriesColor(s);
            var line = new List<(double X, double Y)>();
            for (int i = 0; i < series[s].Values.Count; i++)
            {
                if (series[s].Values[i] is not { } value)
                    continue;

                line.Add((plot.Left + (band * i) + (band / 2), map(value)));
            }

            if (line.Count < 2)
                continue;

            if (filled)
            {
                var polygon = new List<(double, double)>(line) { (line[^1].X, map(0)), (line[0].X, map(0)) };
                page.Items.Add(new PolygonItem { Points = polygon, Color = color });
            }
            else
            {
                for (int i = 1; i < line.Count; i++)
                {
                    page.Items.Add(new StrokeItem
                    {
                        X = line[i - 1].X, Y = line[i - 1].Y, X2 = line[i].X, Y2 = line[i].Y,
                        Thickness = 1.5, Color = color,
                    });
                }
            }
        }

        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    /// <summary>One ring of shares, each wedge a fan of short segments.</summary>
    private void DrawPie(ComposedPage page, SeriesData series, Rect plot, bool doughnut)
    {
        double total = 0;
        foreach (double? value in series.Values)
            total += Math.Abs(value ?? 0);
        if (total <= 0)
            return;

        double radius = Math.Min(plot.Width, plot.Height) / 2;
        double centreX = plot.Left + (plot.Width / 2);
        double centreY = plot.Top + (plot.Height / 2);
        double hole = doughnut ? radius * 0.5 : 0;

        double angle = -Math.PI / 2;    // Twelve o'clock, clockwise — where Excel starts.
        for (int i = 0; i < series.Values.Count; i++)
        {
            double share = Math.Abs(series.Values[i] ?? 0) / total;
            if (share <= 0)
                continue;

            double sweep = share * 2 * Math.PI;
            int steps = Math.Max(2, (int)(sweep / 0.1));

            var wedge = new List<(double, double)>();
            for (int step = 0; step <= steps; step++)
            {
                double a = angle + (sweep * step / steps);
                wedge.Add((centreX + (radius * Math.Cos(a)), centreY + (radius * Math.Sin(a))));
            }

            if (hole > 0)
            {
                for (int step = steps; step >= 0; step--)
                {
                    double a = angle + (sweep * step / steps);
                    wedge.Add((centreX + (hole * Math.Cos(a)), centreY + (hole * Math.Sin(a))));
                }
            }
            else
            {
                wedge.Add((centreX, centreY));
            }

            page.Items.Add(new PolygonItem { Points = wedge, Color = SeriesColor(i) });
            angle += sweep;
        }
    }
}
