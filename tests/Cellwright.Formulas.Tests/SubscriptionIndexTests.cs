using Cellwright.Formulas.Engine;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The subscription index replaces a linear scan, so the only thing that matters is that it
/// answers exactly what the scan answered.
/// </summary>
public sealed class SubscriptionIndexTests
{
    [Fact]
    public void AgreesWithBruteForceOnRandomWorkloads()
    {
        var random = new Random(20260725);
        for (int trial = 0; trial < 200; trial++)
        {
            List<(RangeKey Range, int Entry)> subscriptions = GenerateSubscriptions(random, random.Next(1, 60));
            RangeSubscriptionIndex index = Build(subscriptions);

            for (int probe = 0; probe < 40; probe++)
            {
                int sheetId = random.Next(0, 3);
                int row = random.Next(1, 60);
                int column = random.Next(1, 40);

                var fromIndex = new HashSet<int>();
                index.Collect(sheetId, row, column, fromIndex);

                var expected = new HashSet<int>();
                foreach ((RangeKey range, int entry) in subscriptions)
                {
                    if (range.SheetId == sheetId && range.Contains(row, column))
                        expected.Add(entry);
                }

                Assert.Equal(expected.OrderBy(static x => x), fromIndex.OrderBy(static x => x));
            }
        }
    }

    [Fact]
    public void IdenticalRangesAreMergedIntoOneSubscription()
    {
        var range = new RangeKey(0, 1, 1, 100, 2);
        var builder = new RangeSubscriptionIndex.Builder();
        for (int entry = 0; entry < 5_000; entry++)
            builder.Add(range, entry);

        RangeSubscriptionIndex index = builder.Build();
        Assert.Equal(1, index.SubscriptionCount);

        var hits = new HashSet<int>();
        index.Collect(0, 50, 1, hits);
        Assert.Equal(5_000, hits.Count);
    }

    [Fact]
    public void WholeColumnSubscriptionsAreSeparatedByColumn()
    {
        // The degenerate case for a row-only index: every range spans all rows, so only the
        // column tells them apart.
        var builder = new RangeSubscriptionIndex.Builder();
        for (int column = 1; column <= 500; column++)
            builder.Add(new RangeKey(0, 1, column, 10_000, column), column);

        RangeSubscriptionIndex index = builder.Build();

        var hits = new HashSet<int>();
        index.Collect(0, 5_000, 42, hits);
        Assert.Equal([42], hits);
    }

    [Fact]
    public void RangesWiderThanTheBucketLimitStillMatch()
    {
        int width = RangeSubscriptionIndex.MaxBucketColumns + 10;
        var builder = new RangeSubscriptionIndex.Builder();
        builder.Add(new RangeKey(0, 5, 1, 15, width), 7);
        RangeSubscriptionIndex index = builder.Build();

        var inside = new HashSet<int>();
        index.Collect(0, 10, width, inside);
        Assert.Equal([7], inside);

        var outside = new HashSet<int>();
        index.Collect(0, 10, width + 1, outside);
        Assert.Empty(outside);
        index.Collect(0, 4, 1, outside);
        Assert.Empty(outside);
    }

    [Fact]
    public void NestedAndTouchingIntervalsAllMatch()
    {
        var builder = new RangeSubscriptionIndex.Builder();
        builder.Add(new RangeKey(0, 1, 1, 100, 1), 1);   // Encloses the rest.
        builder.Add(new RangeKey(0, 40, 1, 60, 1), 2);
        builder.Add(new RangeKey(0, 50, 1, 50, 1), 3);   // Single row.
        builder.Add(new RangeKey(0, 61, 1, 70, 1), 4);   // Starts right after 2.
        RangeSubscriptionIndex index = builder.Build();

        var hits = new HashSet<int>();
        index.Collect(0, 50, 1, hits);
        Assert.Equal([1, 2, 3], hits.OrderBy(static x => x));

        hits.Clear();
        index.Collect(0, 61, 1, hits);
        Assert.Equal([1, 4], hits.OrderBy(static x => x));
    }

    [Fact]
    public void OtherSheetsAreNotConfused()
    {
        var builder = new RangeSubscriptionIndex.Builder();
        builder.Add(new RangeKey(0, 1, 1, 10, 1), 100);
        builder.Add(new RangeKey(1, 1, 1, 10, 1), 200);
        RangeSubscriptionIndex index = builder.Build();

        var hits = new HashSet<int>();
        index.Collect(1, 5, 1, hits);
        Assert.Equal([200], hits);
    }

    private static RangeSubscriptionIndex Build(List<(RangeKey Range, int Entry)> subscriptions)
    {
        var builder = new RangeSubscriptionIndex.Builder();
        foreach ((RangeKey range, int entry) in subscriptions)
            builder.Add(range, entry);
        return builder.Build();
    }

    /// <summary>Mixes the shapes real workbooks produce: single columns, small blocks and wide bands.</summary>
    private static List<(RangeKey Range, int Entry)> GenerateSubscriptions(Random random, int count)
    {
        var subscriptions = new List<(RangeKey, int)>(count);
        for (int entry = 0; entry < count; entry++)
        {
            int sheetId = random.Next(0, 3);
            int firstRow = random.Next(1, 50);
            int firstColumn = random.Next(1, 30);
            int height = random.Next(4) switch
            {
                0 => 1,
                1 => random.Next(2, 10),
                2 => random.Next(10, 40),
                _ => 60,
            };
            int width = random.Next(4) switch
            {
                0 => 1,
                1 => random.Next(2, 5),
                2 => RangeSubscriptionIndex.MaxBucketColumns,
                _ => RangeSubscriptionIndex.MaxBucketColumns + random.Next(1, 12),
            };

            subscriptions.Add((new RangeKey(
                sheetId, firstRow, firstColumn, firstRow + height - 1, firstColumn + width - 1), entry));
        }

        return subscriptions;
    }
}
