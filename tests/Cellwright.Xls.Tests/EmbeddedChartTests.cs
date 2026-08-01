using System.Buffers.Binary;
using Cellwright.Xls.Cfb;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// A sheet may hold substreams of its own — an embedded chart is one — and the records after
/// such a substream still belong to the sheet. Reading them requires telling a nested BOF from
/// the one that starts the next sheet, which nothing in the record itself says.
/// </summary>
public sealed class EmbeddedChartTests
{
    /// <summary>Record that opens a substream, and the one that closes it.</summary>
    private const ushort Bof = 0x0809;
    private const ushort Eof = 0x000A;

    /// <summary>Where the writer puts the drawing layer, which is what a chart would displace.</summary>
    private const ushort MsoDrawing = 0x00EC;

    [Fact]
    public async Task EverythingAfterAnEmbeddedChartIsStillRead()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Data");
        sheet["A1"] = CellValue.Text("value");
        sheet.Comments.Add("B2", "after the chart", "Tester");
        sheet.Merge(RangeRef.Parse("D1:E1"));
        sheet.FrozenRows = 1;
        sheet.ShowGridLines = false;

        byte[] withChart = await InjectChartSubstreamAsync(original);

        using var buffer = new MemoryStream(withChart);
        var warnings = new List<WorkbookWarning>();
        Workbook read = await XlsWorkbook.LoadAsync(
            buffer,
            new XlsLoadOptions { LeaveOpen = true, OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        Worksheet loaded = read.Sheets[0];
        Assert.Equal("value", loaded["A1"].GetText());
        Assert.Equal("after the chart", Assert.Single(loaded.Comments).Text);
        Assert.Equal(RangeRef.Parse("D1:E1"), Assert.Single(loaded.MergedRanges));
        Assert.Equal(1, loaded.FrozenRows);
        Assert.False(loaded.ShowGridLines);

        // The sheet ends where it always did, so nothing was recovered from.
        Assert.DoesNotContain(warnings, static warning => warning.Code == "MissingEof");
    }

    /// <summary>
    /// Writes the workbook, then splices an empty chart substream into its sheet just before
    /// the drawing records — which is where Excel puts one.
    /// </summary>
    private static async Task<byte[]> InjectChartSubstreamAsync(Workbook workbook)
    {
        using var plain = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, plain, options: null, TestContext.Current.CancellationToken);

        CompoundFile container = CompoundFile.Parse(plain.ToArray());
        byte[] stream = container.TryReadStream("Workbook")!;
        int insertAt = OffsetOfFirst(stream, MsoDrawing);

        byte[] chart = ChartSubstream();
        var spliced = new byte[stream.Length + chart.Length];
        stream.AsSpan(0, insertAt).CopyTo(spliced);
        chart.CopyTo(spliced.AsSpan(insertAt));
        stream.AsSpan(insertAt).CopyTo(spliced.AsSpan(insertAt + chart.Length));

        using var repacked = new MemoryStream();
        await CompoundFileWriter.WriteAsync(repacked, "Workbook", spliced, TestContext.Current.CancellationToken);
        return repacked.ToArray();
    }

    /// <summary>A chart substream with nothing in it, which is all the nesting needs to be real.</summary>
    private static byte[] ChartSubstream()
    {
        var records = new List<byte>();
        records.AddRange([0x09, 0x08, 0x10, 0x00]);                 // BOF, sixteen bytes of body.
        records.AddRange([0x00, 0x06]);                             // BIFF8.
        records.AddRange([0x20, 0x00]);                             // A chart.
        records.AddRange(new byte[12]);
        records.AddRange([0x0A, 0x00, 0x00, 0x00]);                 // EOF.
        return [.. records];
    }

    /// <summary>Finds where a record of the given type starts, walking the stream by headers.</summary>
    private static int OffsetOfFirst(byte[] stream, ushort recordType)
    {
        int offset = 0;
        while (offset + 4 <= stream.Length)
        {
            ushort type = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset));
            int size = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset + 2));
            if (type == recordType)
                return offset;

            offset += 4 + size;
        }

        throw new InvalidOperationException($"No record of type 0x{recordType:X4} in the stream.");
    }

    /// <summary>
    /// The corpus has real workbooks with embedded charts. None of them should look like a
    /// sheet that lost its end marker, which is what a nested substream used to be taken for.
    /// </summary>
    [Fact]
    public async Task RealWorkbooksWithChartsAreNotTakenForBrokenOnes()
    {
        string[] names =
        [
            "WithChart.xls", "WithTwoCharts.xls", "WithThreeCharts.xls",
            "SimpleChart.xls", "chartdemo.xls", "EmbeddedChartHeaderTest.xls",
        ];

        int examined = 0;
        foreach (string name in names)
        {
            string? path = XlsCorpus.Files.FirstOrDefault(
                file => Path.GetFileName(file).Equals(name, StringComparison.OrdinalIgnoreCase));
            if (path is null)
                continue;

            examined++;
            var warnings = new List<WorkbookWarning>();
            Workbook workbook = await XlsWorkbook.LoadAsync(
                path, new XlsLoadOptions { OnWarning = warnings.Add }, TestContext.Current.CancellationToken);

            Assert.NotEmpty(workbook.Sheets);
            Assert.DoesNotContain(warnings, warning => warning.Code == "MissingEof");
        }

        Assert.SkipWhen(examined == 0, "No workbook with an embedded chart is in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
    }
}
