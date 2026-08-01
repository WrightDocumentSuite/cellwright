using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Reading the charts a workbook carries: the frame's anchor, the kind, the title, and the
/// series with their references and cached points.
/// </summary>
public sealed class ChartReadTests
{
    /// <summary>A workbook holding one sheet with one chart, built part by part.</summary>
    private static MemoryStream BuildChartedWorkbook(string chartSpace)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
                <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
                </Types>
                """);

            Write(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            Write(archive, "xl/workbook.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets><sheet name="Данные" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);

            Write(archive, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """);

            Write(archive, "xl/worksheets/sheet1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c></row></sheetData>
                <drawing r:id="rId1"/>
                </worksheet>
                """);

            Write(archive, "xl/worksheets/_rels/sheet1.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
                </Relationships>
                """);

            Write(archive, "xl/drawings/drawing1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <xdr:twoCellAnchor>
                <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId1"/>
                </a:graphicData></a:graphic>
                </xdr:graphicFrame>
                <xdr:clientData/>
                </xdr:twoCellAnchor>
                </xdr:wsDr>
                """);

            Write(archive, "xl/drawings/_rels/drawing1.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """);

            Write(archive, "xl/charts/chart1.xml", chartSpace);
        }

        stream.Position = 0;
        return stream;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using Stream entry = archive.CreateEntry(name).Open();
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        entry.Write(bytes);
    }

    /// <summary>A column chart with one series, spelled the way Excel spells it.</summary>
    private const string ColumnChart = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart>
        <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Выручка</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:plotArea><c:layout/>
        <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>Данные!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Продажи</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>Данные!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
        <c:pt idx="0"><c:v>Янв</c:v></c:pt><c:pt idx="1"><c:v>Фев</c:v></c:pt><c:pt idx="2"><c:v>Мар</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>Данные!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
        <c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>150.5</c:v></c:pt><c:pt idx="2"><c:v>90</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
        </c:barChart>
        </c:plotArea>
        </c:chart>
        </c:chartSpace>
        """;

    private static async Task<Chart> LoadChartAsync(string chartSpace)
    {
        using MemoryStream source = BuildChartedWorkbook(chartSpace);
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        return Assert.Single(book.Sheets[0].Charts);
    }

    [Fact]
    public async Task TheChartIsFoundThroughItsFrame()
    {
        Chart chart = await LoadChartAsync(ColumnChart);

        Assert.Equal("Chart 1", chart.Name);
        Assert.Equal(PictureAnchorKind.TwoCell, chart.Anchor.Kind);
        Assert.Equal(4, chart.Anchor.From.Column);      // Zero-based 3 in the file.
        Assert.Equal(2, chart.Anchor.From.Row);
    }

    [Fact]
    public async Task TheKindAndTitleAreParsed()
    {
        Chart chart = await LoadChartAsync(ColumnChart);

        Assert.Equal(ChartKind.Column, chart.Kind);
        Assert.Equal("Выручка", chart.Title);
    }

    [Fact]
    public async Task TheSeriesCarriesItsReferencesAndItsCaches()
    {
        Chart chart = await LoadChartAsync(ColumnChart);
        ChartSeries series = Assert.Single(chart.Series);

        Assert.Equal("Продажи", series.Name);
        Assert.Equal("Данные!$B$1", series.NameReference);
        Assert.Equal("Данные!$A$2:$A$4", series.CategoriesReference);
        Assert.Equal("Данные!$B$2:$B$4", series.ValuesReference);
        Assert.Equal(["Янв", "Фев", "Мар"], series.Categories);
        Assert.Equal([100d, 150.5, 90d], series.Values);
    }

    [Fact]
    public async Task ABarDirectionMakesItABarChart()
    {
        Chart chart = await LoadChartAsync(ColumnChart.Replace("val=\"col\"", "val=\"bar\"", StringComparison.Ordinal));

        Assert.Equal(ChartKind.Bar, chart.Kind);
    }

    [Fact]
    public async Task APieChartIsAPieChart()
    {
        // Renames both tags at once; the barDir left inside is ignored by the parser, and its
        // guard keeps it from turning the pie back into a bar.
        Chart chart = await LoadChartAsync(
            ColumnChart.Replace("barChart>", "pieChart>", StringComparison.Ordinal));

        Assert.Equal(ChartKind.Pie, chart.Kind);
    }

    [Fact]
    public async Task TheSeriesReferenceResolvesToLiveCells()
    {
        // The bridge to the grid: the reference the series carries names real cells, which is
        // what a renderer without caches will read.
        using MemoryStream source = BuildChartedWorkbook(ColumnChart);
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Chart chart = Assert.Single(book.Sheets[0].Charts);
        ChartSeries series = Assert.Single(chart.Series);

        (Worksheet sheet, RangeRef range) = SheetRangeResolver.Resolve(book, series.ValuesReference!)!.Value;
        Assert.Equal("Данные", sheet.Name);
        Assert.Equal(RangeRef.Parse("B2:B4"), range);
    }

    [Fact]
    public async Task TheChartSurvivesTheRoundTripUntouched()
    {
        using MemoryStream source = BuildChartedWorkbook(ColumnChart);
        Workbook book = await Workbook.LoadAsync(
            source, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        saved.Position = 0;

        Workbook reloaded = await Workbook.LoadAsync(
            saved, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Chart chart = Assert.Single(reloaded.Sheets[0].Charts);
        Assert.Equal(ChartKind.Column, chart.Kind);
        Assert.Equal("Выручка", chart.Title);
    }
}
