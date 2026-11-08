using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Creating a chart and writing it: the generated chartSpace, the frame in the drawing, and
/// the round trip back through our own reader and the official validator.
/// </summary>
public sealed class ChartWriteTests
{
    private static Worksheet Data(out Workbook book)
    {
        book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Данные");
        sheet["A1"] = "Янв";
        sheet["A2"] = "Фев";
        sheet["A3"] = "Мар";
        sheet["B1"] = 100d;
        sheet["B2"] = 150.5;
        sheet["B3"] = 90d;
        return sheet;
    }

    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task ACreatedChartComesBackThroughOurOwnReader()
    {
        Worksheet sheet = Data(out Workbook book);
        sheet.AddChart(
            ChartKind.Column, "D2:K17", "Данные!$B$1:$B$3",
            categoriesReference: "Данные!$A$1:$A$3", title: "Выручка", seriesName: "Продажи");

        Workbook reloaded = await SaveAndReloadAsync(book);
        Chart chart = Assert.Single(reloaded.Sheets[0].Charts);

        Assert.Equal(ChartKind.Column, chart.Kind);
        Assert.Equal("Выручка", chart.Title);

        ChartSeries series = Assert.Single(chart.Series);
        Assert.Equal("Продажи", series.Name);
        Assert.Equal("Данные!$B$1:$B$3", series.ValuesReference);
        Assert.Equal([100d, 150.5, 90d], series.Values);
        Assert.Equal(["Янв", "Фев", "Мар"], series.Categories);
    }

    [Fact]
    public async Task TheCachesAreFilledFromTheLiveCells()
    {
        Worksheet sheet = Data(out Workbook book);
        Chart chart = sheet.AddChart(ChartKind.Line, "D2:K17", "Данные!$B$1:$B$3");

        // Before any save: the created chart already carries what the cells held.
        Assert.Equal([100d, 150.5, 90d], Assert.Single(chart.Series).Values);
        await Task.CompletedTask;
    }

    [Fact]
    public async Task TheAnchorCoversTheAskedCells()
    {
        Worksheet sheet = Data(out Workbook book);
        sheet.AddChart(ChartKind.Column, "D2:K17", "Данные!$B$1:$B$3");

        Workbook reloaded = await SaveAndReloadAsync(book);
        PictureAnchor anchor = Assert.Single(reloaded.Sheets[0].Charts).Anchor;

        Assert.Equal(CellRef.Parse("D2"), anchor.From);
        // The far corner is exclusive in the drawing, so K17 covered means the anchor reaches L18.
        Assert.Equal(CellRef.Parse("L18"), anchor.To);
    }

    [Fact]
    public async Task AChartAndAPictureShareTheDrawing()
    {
        Worksheet sheet = Data(out Workbook book);
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        sheet.Pictures.Add(png, "A5", "B6");
        sheet.AddChart(ChartKind.Pie, "D2:K17", "Данные!$B$1:$B$3");

        Workbook reloaded = await SaveAndReloadAsync(book);

        Assert.Single(reloaded.Sheets[0].Pictures);
        Chart chart = Assert.Single(reloaded.Sheets[0].Charts);
        Assert.Equal(ChartKind.Pie, chart.Kind);
    }

    [Fact]
    public async Task EveryCreatableKindSurvivesTheRoundTrip()
    {
        foreach (ChartKind kind in new[]
        {
            ChartKind.Column, ChartKind.Bar, ChartKind.Line,
            ChartKind.Pie, ChartKind.Doughnut, ChartKind.Area,
            ChartKind.Scatter, ChartKind.Bubble, ChartKind.Radar,
        })
        {
            Worksheet sheet = Data(out Workbook book);
            sheet.AddChart(kind, "D2:K17", "Данные!$B$1:$B$3", categoriesReference: "Данные!$A$1:$A$3");

            Workbook reloaded = await SaveAndReloadAsync(book);

            Assert.Equal(kind, Assert.Single(reloaded.Sheets[0].Charts).Kind);
        }
    }

    [Fact]
    public void AKindThatCannotBeCreatedSaysSo()
    {
        Worksheet sheet = Data(out _);

        Assert.Throws<ArgumentException>(() =>
            sheet.AddChart(ChartKind.Unknown, "D2:K17", "Данные!$B$1:$B$3"));
    }

    [Fact]
    public async Task MutatingChartTitleAndSeriesUpdatesCachesAndSurvivesRoundTrip()
    {
        Worksheet sheet = Data(out Workbook book);
        Chart chart = sheet.AddChart(ChartKind.Column, "D2:K17", "Данные!$B$1:$B$2", title: "Старый");

        sheet["B3"] = 250d;
        chart.SetTitle("Новый заголовок");
        chart.SetData("Данные!$B$1:$B$3", "Данные!$A$1:$A$3", "Обновлено");

        Assert.Equal("Новый заголовок", chart.Title);
        Assert.Equal([100d, 150.5, 250d], Assert.Single(chart.Series).Values);

        Workbook reloaded = await SaveAndReloadAsync(book);
        Chart reloadedChart = Assert.Single(reloaded.Sheets[0].Charts);
        Assert.Equal("Новый заголовок", reloadedChart.Title);
        Assert.Equal("Обновлено", Assert.Single(reloadedChart.Series).Name);
        Assert.Equal([100d, 150.5, 250d], Assert.Single(reloadedChart.Series).Values);
    }

    [Fact]
    public async Task TwoSeriesPlotInOrder()
    {
        Worksheet sheet = Data(out Workbook book);
        sheet["C1"] = 10d;
        sheet["C2"] = 20d;
        sheet["C3"] = 30d;

        sheet.AddChart(ChartKind.Column, "D2:K17", [
            new ChartSeriesDefinition("Данные!$B$1:$B$3", Name: "Первая"),
            new ChartSeriesDefinition("Данные!$C$1:$C$3", Name: "Вторая"),
        ]);

        Workbook reloaded = await SaveAndReloadAsync(book);
        Chart chart = Assert.Single(reloaded.Sheets[0].Charts);

        Assert.Equal(2, chart.Series.Count);
        Assert.Equal("Первая", chart.Series[0].Name);
        Assert.Equal([10d, 20d, 30d], chart.Series[1].Values);
    }

    [Fact]
    public async Task ALoadedChartSurvivesBesideACreatedOne()
    {
        // Adding a chart regenerates the drawing; the chart that came from the file must ride
        // along rather than be lost with the old part.
        Worksheet sheet = Data(out Workbook book);
        sheet.AddChart(ChartKind.Column, "D2:K17", "Данные!$B$1:$B$3", title: "Первая");

        Workbook once = await SaveAndReloadAsync(book);
        once.Sheets[0].AddChart(ChartKind.Line, "D20:K35", "Данные!$B$1:$B$3", title: "Вторая");

        Workbook twice = await SaveAndReloadAsync(once);

        Assert.Equal(2, twice.Sheets[0].Charts.Count);
        Assert.Contains(twice.Sheets[0].Charts, chart => chart.Title == "Первая");
        Assert.Contains(twice.Sheets[0].Charts, chart => chart.Title == "Вторая");
    }

    [Fact]
    public async Task InsertingRowsAndColumnsShiftsChartAnchorAndSeriesReferences()
    {
        Worksheet sheet = Data(out Workbook book);
        Chart chart = sheet.AddChart(
            ChartKind.Column, "D2:K17", "Данные!$B$1:$B$3",
            categoriesReference: "Данные!$A$1:$A$3", title: "Динамика");

        sheet.InsertRows(1, 2);
        sheet.InsertColumns(1, 1);

        Assert.Equal(CellRef.Parse("E4"), chart.Anchor.From);
        Assert.Equal(CellRef.Parse("M20"), chart.Anchor.To);

        ChartSeries liveSeries = Assert.Single(chart.Series);
        Assert.Equal("Данные!$C$3:$C$5", liveSeries.ValuesReference);
        Assert.Equal("Данные!$B$3:$B$5", liveSeries.CategoriesReference);

        Workbook reloaded = await SaveAndReloadAsync(book);
        Chart reloadedChart = Assert.Single(reloaded.Sheets[0].Charts);
        Assert.Equal(CellRef.Parse("E4"), reloadedChart.Anchor.From);
        Assert.Equal(CellRef.Parse("M20"), reloadedChart.Anchor.To);

        ChartSeries reloadedSeries = Assert.Single(reloadedChart.Series);
        Assert.Equal("Данные!$C$3:$C$5", reloadedSeries.ValuesReference);
        Assert.Equal("Данные!$B$3:$B$5", reloadedSeries.CategoriesReference);
    }

    [Fact]
    public async Task ABubbleChartAndCombinedChartSurviveRoundTripAndValidation()
    {
        Worksheet sheet = Data(out Workbook book);
        sheet["C1"] = 10d;
        sheet["C2"] = 30d;
        sheet["C3"] = 20d;

        sheet.AddChart(ChartKind.Bubble, "D2:K17", [
            new ChartSeriesDefinition(
                ValuesReference: "Данные!$B$1:$B$3",
                CategoriesReference: "Данные!$A$1:$A$3",
                Name: "Пузырьки",
                BubbleSizesReference: "Данные!$C$1:$C$3"),
        ], title: "Bubble");

        sheet.AddChart(ChartKind.Column, "D20:K35", [
            new ChartSeriesDefinition("Данные!$B$1:$B$3", "Данные!$A$1:$A$3", "Столбцы", Kind: ChartKind.Column),
            new ChartSeriesDefinition("Данные!$C$1:$C$3", "Данные!$A$1:$A$3", "Линия", Kind: ChartKind.Line),
        ], title: "Комбинированная");

        Workbook reloaded = await SaveAndReloadAsync(book);
        Assert.Equal(2, reloaded.Sheets[0].Charts.Count);

        Chart bubble = reloaded.Sheets[0].Charts[0];
        Assert.Equal(ChartKind.Bubble, bubble.Kind);
        ChartSeries bubbleSeries = Assert.Single(bubble.Series);
        Assert.Equal(ChartKind.Bubble, bubbleSeries.Kind);
        Assert.Equal("Данные!$C$1:$C$3", bubbleSeries.BubbleSizesReference);
        Assert.Equal([10d, 30d, 20d], bubbleSeries.BubbleSizes);

        Chart combined = reloaded.Sheets[0].Charts[1];
        Assert.Equal(ChartKind.Column, combined.Kind);
        Assert.Equal(2, combined.Series.Count);
        Assert.Equal(ChartKind.Column, combined.Series[0].Kind);
        Assert.Equal(ChartKind.Line, combined.Series[1].Kind);
    }
}
