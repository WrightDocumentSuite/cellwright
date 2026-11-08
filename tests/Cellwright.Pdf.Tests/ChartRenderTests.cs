using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Charts on the printed page: the title and labels that land as text, the marks that land as
/// fills and paths, and the charts that are honestly reported instead of half-drawn.
/// </summary>
public sealed class ChartRenderTests
{
    private static Workbook Charted(ChartKind kind, string? title = "Выручка")
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Данные");
        sheet["A1"] = "Янв";
        sheet["A2"] = "Фев";
        sheet["A3"] = "Мар";
        sheet["B1"] = 100d;
        sheet["B2"] = 150d;
        sheet["B3"] = 90d;

        sheet.AddChart(kind, "D2:K17", "Данные!$B$1:$B$3",
            categoriesReference: "Данные!$A$1:$A$3", title: title, seriesName: "Продажи");

        return book;
    }

    [Fact]
    public void TheTitleIsDrawn()
    {
        using Rendered rendered = Rendered.Of(Charted(ChartKind.Column));

        Assert.Contains("Выручка", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheCategoryLabelsAreDrawn()
    {
        using Rendered rendered = Rendered.Of(Charted(ChartKind.Column));

        Assert.Contains("Янв", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("Мар", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheValueAxisLandsOnRoundNumbers()
    {
        using Rendered rendered = Rendered.Of(Charted(ChartKind.Column));

        // Values run 0..150, so a round-numbered axis shows 0, 50, 100, 150.
        Assert.Contains("150", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("50", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void APieIsDrawnFromWedges()
    {
        using Rendered rendered = Rendered.Of(Charted(ChartKind.Pie, title: null));

        // Wedges reach the page as paths — fans of line segments — rather than rectangles.
        Assert.Contains(" l\n", rendered.Content(), StringComparison.Ordinal);

        // And its legend names the slices.
        Assert.Contains("Фев", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void EveryCreatableKindRendersWithoutComplaint()
    {
        foreach (ChartKind kind in new[]
        {
            ChartKind.Column, ChartKind.Bar, ChartKind.Line,
            ChartKind.Pie, ChartKind.Doughnut, ChartKind.Area,
            ChartKind.Scatter, ChartKind.Bubble, ChartKind.Radar,
        })
        {
            using Rendered rendered = Rendered.Of(Charted(kind));

            Assert.DoesNotContain(
                rendered.Diagnostics,
                warning => warning.Subject?.StartsWith("chart", StringComparison.Ordinal) == true);
        }
    }

    [Fact]
    public void ARadarChartDrawsSpokesWebAndCategoryLabels()
    {
        using Rendered rendered = Rendered.Of(Charted(ChartKind.Radar));

        Assert.Contains("Выручка", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("Янв", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("S\n", rendered.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void AScatterChartPlotsNumericCoordinateMarkers()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("XY");
        sheet["A1"] = 10d;
        sheet["A2"] = 20d;
        sheet["A3"] = 30d;
        sheet["B1"] = 5d;
        sheet["B2"] = 15d;
        sheet["B3"] = 25d;
        sheet.AddChart(ChartKind.Scatter, "D2:K17", "XY!$B$1:$B$3", categoriesReference: "XY!$A$1:$A$3", title: "Scatter Plot");

        using Rendered rendered = Rendered.Of(book);
        Assert.Contains("Scatter Plot", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("re\nf", rendered.Content(), StringComparison.Ordinal);
    }

    [Fact]
    public void ABubbleChartAndCombinedChartRenderWithoutWarnings()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Данные");
        sheet["A1"] = "Янв";
        sheet["A2"] = "Фев";
        sheet["A3"] = "Мар";
        sheet["B1"] = 100d;
        sheet["B2"] = 150d;
        sheet["B3"] = 90d;
        sheet["C1"] = 10d;
        sheet["C2"] = 30d;
        sheet["C3"] = 20d;

        sheet.AddChart(ChartKind.Bubble, "D2:K17", [
            new ChartSeriesDefinition("Данные!$B$1:$B$3", "Данные!$A$1:$A$3", "Пузырьки", "Данные!$C$1:$C$3"),
        ], title: "Bubble Chart");

        sheet.AddChart(ChartKind.Column, "D19:K34", [
            new ChartSeriesDefinition("Данные!$B$1:$B$3", "Данные!$A$1:$A$3", "Факт", Kind: ChartKind.Column),
            new ChartSeriesDefinition("Данные!$C$1:$C$3", "Данные!$A$1:$A$3", "Тренд", Kind: ChartKind.Line),
        ], title: "Combo Chart");

        using Rendered rendered = Rendered.Of(book);
        Assert.Contains("Bubble Chart", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("Combo Chart", rendered.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("chart", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void AChartAcrossAPageBreakIsReportedNotHalfDrawn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Данные");
        for (int i = 1; i <= 3; i++)
            sheet[$"B{i}"] = i * 10d;

        // Far off to the right: the anchor runs past what one page holds.
        sheet.AddChart(ChartKind.Column, "A1:AZ40", "Данные!$B$1:$B$3");

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "chart-split");
    }

    [Fact]
    public async Task ALoadedChartRendersFromItsCaches()
    {
        // Through a save and reload: the chart renders from what the file carries, the same
        // path a chart from Excel takes.
        Workbook book = Charted(ChartKind.Column);

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using Rendered rendered = Rendered.Of(reloaded);

        Assert.Contains("Выручка", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("Янв", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void ASheetWithoutChartsSaysNothingAboutThem()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = "plain";

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("chart", StringComparison.Ordinal) == true);
    }
}
