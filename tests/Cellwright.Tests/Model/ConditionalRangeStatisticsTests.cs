using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>Regression and allocation checks for one-pass conditional-range aggregation.</summary>
public sealed class ConditionalRangeStatisticsTests
{
    [Fact]
    public void DenseDisjointRangesDoNotPayForACoordinateSet()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int row = 1; row <= 8_192; row++)
            sheet[row, 1] = row;

        RangeRef[] contiguous = [RangeRef.Parse("A1:A8192")];
        RangeRef[] disjoint = [RangeRef.Parse("A1:A4096"), RangeRef.Parse("A4097:A8192")];

        // Warm the iterator and sorting paths before measuring the current test thread.
        _ = Gather(sheet, contiguous, book.DateSystem);
        _ = Gather(sheet, disjoint, book.DateSystem);

        long contiguousBytes = AllocatedBy(() => Gather(sheet, contiguous, book.DateSystem));
        long disjointBytes = AllocatedBy(() => Gather(sheet, disjoint, book.DateSystem));

        // A second range needs only its iterator objects. A HashSet<CellRef> for all 8,192
        // coordinates would make this delta hundreds of kilobytes instead of a few hundred.
        Assert.InRange(disjointBytes - contiguousBytes, 0, 512);
    }

    [Fact]
    public void OverlappingRangesStillCountEachCoordinateOnlyOnce()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int row = 1; row <= 5; row++)
            sheet[row, 1] = row;

        ConditionalRangeStatistics statistics = ConditionalRangeStatistics.Gather(
            sheet,
            [RangeRef.Parse("A1:A3"), RangeRef.Parse("A3:A5")],
            countOccurrences: true,
            book.DateSystem);

        Assert.Equal(5, statistics.Count);
        Assert.Equal(15, statistics.Sum);
        Assert.Equal(6, statistics.OccupiedSlotsInspected);
        Assert.Equal(1, statistics.OccurrencesOf(sheet[3, 1]));
    }

    [Fact]
    public void ThousandsOfDisjointSparseAreasInspectEachOccupiedCellOnce()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        const int Areas = 4_096;
        var ranges = new RangeRef[Areas];
        for (int index = 0; index < Areas; index++)
        {
            int row = index + 1;
            sheet[row, 1] = row;
            ranges[index] = new RangeRef(new CellRef(row, 1), new CellRef(row, 1));
        }

        ConditionalRangeStatistics statistics = Gather(sheet, ranges, book.DateSystem);

        Assert.Equal(Areas, statistics.Count);
        Assert.Equal(Areas, statistics.OccupiedSlotsInspected);
    }

    private static ConditionalRangeStatistics Gather(
        Worksheet sheet, IReadOnlyList<RangeRef> ranges, ExcelDateSystem system) =>
        ConditionalRangeStatistics.Gather(sheet, ranges, countOccurrences: false, system);

    private static long AllocatedBy(Func<ConditionalRangeStatistics> gather)
    {
        long before = GC.GetAllocatedBytesForCurrentThread();
        _ = gather();
        return GC.GetAllocatedBytesForCurrentThread() - before;
    }
}
