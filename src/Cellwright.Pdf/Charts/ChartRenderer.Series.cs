using System.Globalization;
using Cellwright.Pdf.Layout;
using Inkwright;

namespace Cellwright.Pdf.Charts;

internal sealed partial class ChartRenderer
{
    private (Rect Plot, Func<double, double> Map) ValueAxis(
        ComposedPage page, List<SeriesData> series, Rect area, bool horizontal)
    {
        (double low, double high) = Extent(series);
        double step = NiceStep(high - low);
        low = Math.Floor(low / step) * step;
        high = Math.Ceiling(high / step) * step;

        double labelRoom = 0;
        for (double tick = low; tick <= high + (step / 2); tick += step)
            labelRoom = Math.Max(labelRoom, _label.Measure(Format(tick)));

        Rect plot = horizontal
            ? new Rect(area.Left, area.Top, area.Right, area.Bottom - (_label.LineHeight * 2) - 2)
            : new Rect(area.Left + labelRoom + 4, area.Top, area.Right, area.Bottom - _label.LineHeight - 2);

        double span = high - low;
        Func<double, double> map = horizontal
            ? value => plot.Left + ((value - low) / span * plot.Width)
            : value => plot.Bottom - ((value - low) / span * plot.Height);

        for (double tick = low; tick <= high + (step / 2); tick += step)
            DrawAxisTick(page, plot, tick, map(tick), horizontal);

        page.Items.Add(new StrokeItem
        {
            X = plot.Left,
            Y = plot.Bottom,
            X2 = plot.Right,
            Y2 = plot.Bottom,
            Thickness = 0.8,
            Color = AxisColor,
        });
        page.Items.Add(new StrokeItem
        {
            X = plot.Left,
            Y = plot.Top,
            X2 = plot.Left,
            Y2 = plot.Bottom,
            Thickness = 0.8,
            Color = AxisColor,
        });

        return (plot, map);
    }

    private void DrawAxisTick(ComposedPage page, Rect plot, double tick, double at, bool horizontal)
    {
        string label = Format(tick);
        double width = _label.Measure(label);
        if (horizontal)
        {
            page.Items.Add(new StrokeItem { X = at, Y = plot.Top, X2 = at, Y2 = plot.Bottom, Thickness = 0.4, Color = GridColor });
            page.Items.Add(new TextItem
            {
                X = at - (width / 2),
                Y = plot.Bottom + 2,
                Text = label,
                Appearance = _label,
                Width = width,
                Height = _label.LineHeight,
            });
            return;
        }

        page.Items.Add(new StrokeItem { X = plot.Left, Y = at, X2 = plot.Right, Y2 = at, Thickness = 0.4, Color = GridColor });
        page.Items.Add(new TextItem
        {
            X = plot.Left - 4 - width,
            Y = at - (_label.LineHeight / 2),
            Text = label,
            Appearance = _label,
            Width = width,
            Height = _label.LineHeight,
        });
    }

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
        int points = MaxPoints(series);
        if (points == 0)
            return;

        double baseline = map(0);
        if (horizontal)
        {
            DrawHorizontalBars(page, series, plot, points, baseline, map);
            DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2 + _label.LineHeight);
            return;
        }

        DrawVerticalBars(page, series, plot, points, baseline, map);
        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    private void DrawHorizontalBars(
        ComposedPage page, List<SeriesData> series, Rect plot, int points, double baseline, Func<double, double> map)
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
    }

    private void DrawVerticalBars(
        ComposedPage page, List<SeriesData> series, Rect plot, int points, double baseline, Func<double, double> map)
    {
        double band = plot.Width / points;
        double bar = band * 0.7 / series.Count;
        for (int s = 0; s < series.Count; s++)
        {
            PdfColor color = SeriesColor(series[s].Index);
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
    }

    private void DrawLines(ComposedPage page, List<SeriesData> series, Rect area, bool filled)
    {
        (Rect plot, Func<double, double> map) = ValueAxis(page, series, area, horizontal: false);
        int points = MaxPoints(series);
        if (points == 0)
            return;

        DrawLineSeries(page, series, plot, points, map, filled);
        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    private void DrawLineSeries(
        ComposedPage page, List<SeriesData> series, Rect plot, int points, Func<double, double> map, bool filled)
    {
        double band = plot.Width / points;
        for (int s = 0; s < series.Count; s++)
        {
            PdfColor color = SeriesColor(series[s].Index);
            var line = new List<(double X, double Y)>();
            for (int i = 0; i < series[s].Values.Count; i++)
            {
                if (series[s].Values[i] is { } value)
                    line.Add((plot.Left + (band * i) + (band / 2), map(value)));
            }

            if (line.Count < 2)
                continue;

            if (filled)
            {
                var polygon = new List<(double, double)>(line) { (line[^1].X, map(0)), (line[0].X, map(0)) };
                page.Items.Add(new PolygonItem { Points = polygon, Color = color });
            }

            StrokePolyline(page, line, color, closed: false);
        }
    }

    private void DrawScatter(ComposedPage page, List<SeriesData> series, Rect area)
    {
        (Rect plot, Func<double, double> mapY) = ValueAxis(page, series, area, horizontal: false);
        int points = MaxPoints(series);
        if (points == 0)
            return;

        DrawScatterSeries(page, series, plot, points, mapY);
        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    private void DrawScatterSeries(
        ComposedPage page, List<SeriesData> series, Rect plot, int points, Func<double, double> mapY)
    {
        bool numericX = TryGetNumericXExtent(series, out double minX, out double maxX);
        double band = plot.Width / points;

        for (int s = 0; s < series.Count; s++)
        {
            PdfColor color = SeriesColor(series[s].Index);
            for (int i = 0; i < series[s].Values.Count; i++)
            {
                if (series[s].Values[i] is not { } yVal)
                    continue;

                double px = numericX && i < series[s].Categories.Count &&
                    double.TryParse(series[s].Categories[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double xVal)
                    ? plot.Left + ((xVal - minX) / Math.Max(1e-9, maxX - minX) * plot.Width)
                    : plot.Left + (band * i) + (band / 2);
                double py = mapY(yVal);
                page.Items.Add(new FillItem { X = px - 2, Y = py - 2, Width = 4, Height = 4, Color = color });
            }
        }
    }

    private void DrawBubble(ComposedPage page, List<SeriesData> series, Rect area)
    {
        (Rect plot, Func<double, double> mapY) = ValueAxis(page, series, area, horizontal: false);
        int points = MaxPoints(series);
        if (points == 0)
            return;

        DrawBubbleSeries(page, series, plot, points, mapY);
        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    private void DrawBubbleSeries(
        ComposedPage page, List<SeriesData> series, Rect plot, int points, Func<double, double> mapY)
    {
        bool numericX = TryGetNumericXExtent(series, out double minX, out double maxX);
        double maxBubble = MaxBubbleSize(series);
        double band = plot.Width / points;

        for (int s = 0; s < series.Count; s++)
        {
            PdfColor color = SeriesColor(series[s].Index);
            for (int i = 0; i < series[s].Values.Count; i++)
            {
                if (series[s].Values[i] is not { } yVal)
                    continue;

                double px = numericX && i < series[s].Categories.Count &&
                    double.TryParse(series[s].Categories[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double xVal)
                    ? plot.Left + ((xVal - minX) / Math.Max(1e-9, maxX - minX) * plot.Width)
                    : plot.Left + (band * i) + (band / 2);
                double py = mapY(yVal);
                double rawSize = i < series[s].BubbleSizes.Count && series[s].BubbleSizes[i] is { } bSize
                    ? Math.Abs(bSize)
                    : Math.Abs(yVal);
                double ratio = maxBubble > 0 ? Math.Sqrt(rawSize / maxBubble) : 0.5;
                double radius = 3.0 + (Math.Clamp(ratio, 0, 1) * 9.0);
                page.Items.Add(new PolygonItem
                {
                    Points = BuildWedge(px, py, radius, 0, 0, 2 * Math.PI),
                    Color = color,
                });
            }
        }
    }

    private void DrawCombinedCartesian(ComposedPage page, List<SeriesData> series, Rect area)
    {
        (Rect plot, Func<double, double> map) = ValueAxis(page, series, area, horizontal: false);
        int points = MaxPoints(series);
        if (points == 0)
            return;

        double baseline = map(0);
        List<SeriesData> areas = [.. series.Where(static s => s.Kind == ChartKind.Area)];
        if (areas.Count > 0)
            DrawLineSeries(page, areas, plot, points, map, filled: true);

        List<SeriesData> columns = [.. series.Where(static s => s.Kind == ChartKind.Column)];
        if (columns.Count > 0)
            DrawVerticalBars(page, columns, plot, points, baseline, map);

        List<SeriesData> lines = [.. series.Where(static s => s.Kind == ChartKind.Line)];
        if (lines.Count > 0)
            DrawLineSeries(page, lines, plot, points, map, filled: false);

        List<SeriesData> scatters = [.. series.Where(static s => s.Kind == ChartKind.Scatter)];
        if (scatters.Count > 0)
            DrawScatterSeries(page, scatters, plot, points, map);

        List<SeriesData> bubbles = [.. series.Where(static s => s.Kind == ChartKind.Bubble)];
        if (bubbles.Count > 0)
            DrawBubbleSeries(page, bubbles, plot, points, map);

        DrawCategoryLabels(page, series[0].Categories, plot, points, plot.Bottom + 2);
    }

    private static double MaxBubbleSize(List<SeriesData> series)
    {
        double max = 0;
        foreach (SeriesData s in series)
        {
            foreach (double? b in s.BubbleSizes)
            {
                if (b is { } value)
                    max = Math.Max(max, Math.Abs(value));
            }
            if (s.BubbleSizes.Count == 0)
            {
                foreach (double? v in s.Values)
                {
                    if (v is { } value)
                        max = Math.Max(max, Math.Abs(value));
                }
            }
        }

        return max;
    }

    private void DrawRadar(ComposedPage page, List<SeriesData> series, Rect plot)
    {
        int spokes = Math.Max(3, MaxPoints(series));
        (double low, double high) = Extent(series);
        double span = Math.Max(1e-9, high - low);
        double cx = plot.Left + (plot.Width / 2);
        double cy = plot.Top + (plot.Height / 2);
        double radius = Math.Max(6, (Math.Min(plot.Width, plot.Height) / 2) - _label.LineHeight - 4);

        DrawRadarWeb(page, series[0].Categories, spokes, cx, cy, radius);
        for (int s = 0; s < series.Count; s++)
        {
            var ring = new List<(double X, double Y)>(spokes);
            for (int i = 0; i < spokes && i < series[s].Values.Count; i++)
            {
                double val = series[s].Values[i] ?? 0;
                double r = Math.Clamp((val - low) / span, 0, 1) * radius;
                double angle = (-Math.PI / 2) + (2 * Math.PI * i / spokes);
                ring.Add((cx + (r * Math.Cos(angle)), cy + (r * Math.Sin(angle))));
            }

            if (ring.Count >= 2)
                StrokePolyline(page, ring, SeriesColor(s), closed: true);
        }
    }

    private void DrawRadarWeb(
        ComposedPage page, IReadOnlyList<string> categories, int spokes, double cx, double cy, double radius)
    {
        for (int level = 1; level <= 4; level++)
        {
            double r = radius * level / 4.0;
            var ring = new List<(double X, double Y)>(spokes);
            for (int i = 0; i < spokes; i++)
            {
                double angle = (-Math.PI / 2) + (2 * Math.PI * i / spokes);
                ring.Add((cx + (r * Math.Cos(angle)), cy + (r * Math.Sin(angle))));
            }

            StrokePolyline(page, ring, GridColor, closed: true, thickness: 0.4);
        }

        for (int i = 0; i < spokes; i++)
        {
            double angle = (-Math.PI / 2) + (2 * Math.PI * i / spokes);
            double ex = cx + (radius * Math.Cos(angle));
            double ey = cy + (radius * Math.Sin(angle));
            page.Items.Add(new StrokeItem { X = cx, Y = cy, X2 = ex, Y2 = ey, Thickness = 0.5, Color = AxisColor });

            if (i < categories.Count && categories[i] is { Length: > 0 } label)
            {
                double w = _label.Measure(label);
                double lx = cx + ((radius + 6) * Math.Cos(angle)) - (w / 2);
                double ly = cy + ((radius + 6) * Math.Sin(angle)) - (_label.LineHeight / 2);
                page.Items.Add(new TextItem { X = lx, Y = ly, Text = label, Appearance = _label, Width = w, Height = _label.LineHeight });
            }
        }
    }

    private static void StrokePolyline(
        ComposedPage page, List<(double X, double Y)> points, PdfColor color, bool closed, double thickness = 1.5)
    {
        for (int i = 1; i < points.Count; i++)
        {
            page.Items.Add(new StrokeItem
            {
                X = points[i - 1].X,
                Y = points[i - 1].Y,
                X2 = points[i].X,
                Y2 = points[i].Y,
                Thickness = thickness,
                Color = color,
            });
        }

        if (closed && points.Count > 2)
        {
            page.Items.Add(new StrokeItem
            {
                X = points[^1].X,
                Y = points[^1].Y,
                X2 = points[0].X,
                Y2 = points[0].Y,
                Thickness = thickness,
                Color = color,
            });
        }
    }

    private static bool TryGetNumericXExtent(List<SeriesData> series, out double minX, out double maxX)
    {
        minX = double.PositiveInfinity;
        maxX = double.NegativeInfinity;
        int parsedCount = 0;
        foreach (SeriesData s in series)
        {
            foreach (string cat in s.Categories)
            {
                if (!double.TryParse(cat, NumberStyles.Float, CultureInfo.InvariantCulture, out double x))
                    return false;
                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                parsedCount++;
            }
        }

        return parsedCount > 0;
    }

    private static int MaxPoints(List<SeriesData> series)
    {
        int points = 0;
        foreach (SeriesData data in series)
            points = Math.Max(points, data.Values.Count);
        return points;
    }

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
        double angle = -Math.PI / 2;

        for (int i = 0; i < series.Values.Count; i++)
        {
            double share = Math.Abs(series.Values[i] ?? 0) / total;
            if (share <= 0)
                continue;

            double sweep = share * 2 * Math.PI;
            page.Items.Add(new PolygonItem
            {
                Points = BuildWedge(centreX, centreY, radius, hole, angle, sweep),
                Color = SeriesColor(i),
            });
            angle += sweep;
        }
    }

    private static List<(double, double)> BuildWedge(
        double cx, double cy, double radius, double hole, double angle, double sweep)
    {
        int steps = Math.Max(2, (int)(sweep / 0.1));
        var wedge = new List<(double, double)>();
        for (int step = 0; step <= steps; step++)
        {
            double a = angle + (sweep * step / steps);
            wedge.Add((cx + (radius * Math.Cos(a)), cy + (radius * Math.Sin(a))));
        }

        if (hole <= 0)
        {
            wedge.Add((cx, cy));
            return wedge;
        }

        for (int step = steps; step >= 0; step--)
        {
            double a = angle + (sweep * step / steps);
            wedge.Add((cx + (hole * Math.Cos(a)), cy + (hole * Math.Sin(a))));
        }

        return wedge;
    }
}
