namespace Cellwright.Formulas.Engine;

/// <summary>
/// Answers "which formulas read this cell through a range reference" without walking every
/// subscription. Ranges are grouped by column, and each column holds an interval tree over
/// rows.
/// </summary>
/// <remarks>
/// Indexing by rows alone does not work: whole-column references are clamped to the used
/// range, so a workbook full of <c>SUM(A:A)</c>-style formulas ends up with thousands of
/// subscriptions all covering rows 1..N. The column is what separates them, and the row tree
/// then separates the partial ranges such as the running totals <c>SUM(A$1:A5)</c>.
/// </remarks>
internal sealed class RangeSubscriptionIndex
{
    /// <summary>
    /// Ranges wider than this are kept in a linear list instead of being copied into every
    /// column bucket. Wide ranges are rare, and duplicating one across hundreds of columns
    /// would cost more than scanning a short list.
    /// </summary>
    public const int MaxBucketColumns = 16;

    private readonly Dictionary<int, SheetIndex> _bySheet;
    private readonly List<int>[] _payloads;

    private RangeSubscriptionIndex(Dictionary<int, SheetIndex> bySheet, List<int>[] payloads)
    {
        _bySheet = bySheet;
        _payloads = payloads;
    }

    /// <summary>Number of distinct ranges tracked, after identical ones were merged.</summary>
    public int SubscriptionCount => _payloads.Length;

    /// <summary>Adds every formula subscribed to a range covering the given cell.</summary>
    public void Collect(int sheetId, int row, int column, HashSet<int> destination)
    {
        if (!_bySheet.TryGetValue(sheetId, out SheetIndex? sheet))
            return;

        if (sheet.Columns.TryGetValue(column, out IntervalTree? tree))
            tree.Stab(row, _payloads, destination);

        foreach (Subscription wide in sheet.Wide)
        {
            if (wide.Contains(row, column))
                AddPayload(wide.Payload, destination);
        }
    }

    private void AddPayload(int payload, HashSet<int> destination)
    {
        foreach (int entry in _payloads[payload])
            destination.Add(entry);
    }

    /// <summary>A range and the formulas that read it.</summary>
    private readonly record struct Subscription(int FirstRow, int LastRow, int FirstColumn, int LastColumn, int Payload)
    {
        public bool Contains(int row, int column) =>
            row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
    }

    private sealed class SheetIndex
    {
        public Dictionary<int, IntervalTree> Columns { get; } = [];

        public List<Subscription> Wide { get; } = [];
    }

    /// <summary>
    /// Static centred interval tree. A stabbing query descends one branch per level, so it
    /// needs neither recursion nor a stack.
    /// </summary>
    private sealed class IntervalTree
    {
        private readonly Node? _root;

        private IntervalTree(Node? root) => _root = root;

        public static IntervalTree Build(List<Subscription> subscriptions) => new(BuildNode(subscriptions));

        public void Stab(int row, List<int>[] payloads, HashSet<int> destination)
        {
            Node? node = _root;
            while (node is not null)
            {
                if (row < node.Center)
                {
                    // Sorted by start: everything after the first interval starting past the
                    // row cannot contain it.
                    foreach (Subscription subscription in node.ByStart)
                    {
                        if (subscription.FirstRow > row)
                            break;
                        Emit(subscription, payloads, destination);
                    }

                    node = node.Left;
                }
                else if (row > node.Center)
                {
                    foreach (Subscription subscription in node.ByEnd)
                    {
                        if (subscription.LastRow < row)
                            break;
                        Emit(subscription, payloads, destination);
                    }

                    node = node.Right;
                }
                else
                {
                    foreach (Subscription subscription in node.ByStart)
                        Emit(subscription, payloads, destination);
                    return;
                }
            }
        }

        private static void Emit(in Subscription subscription, List<int>[] payloads, HashSet<int> destination)
        {
            foreach (int entry in payloads[subscription.Payload])
                destination.Add(entry);
        }

        private static Node? BuildNode(List<Subscription> subscriptions)
        {
            if (subscriptions.Count == 0)
                return null;

            int center = Median(subscriptions);
            List<Subscription> overlapping = [];
            List<Subscription> left = [];
            List<Subscription> right = [];
            foreach (Subscription subscription in subscriptions)
            {
                if (subscription.LastRow < center)
                    left.Add(subscription);
                else if (subscription.FirstRow > center)
                    right.Add(subscription);
                else
                    overlapping.Add(subscription);
            }

            Subscription[] byStart = [.. overlapping];
            Subscription[] byEnd = [.. overlapping];
            Array.Sort(byStart, static (a, b) => a.FirstRow.CompareTo(b.FirstRow));
            Array.Sort(byEnd, static (a, b) => b.LastRow.CompareTo(a.LastRow));

            return new Node(center, byStart, byEnd, BuildNode(left), BuildNode(right));
        }

        private static int Median(List<Subscription> subscriptions)
        {
            // The midpoint of the covered span splits well enough and costs nothing to find,
            // unlike sorting every endpoint at each level.
            int low = int.MaxValue;
            int high = int.MinValue;
            foreach (Subscription subscription in subscriptions)
            {
                low = Math.Min(low, subscription.FirstRow);
                high = Math.Max(high, subscription.LastRow);
            }

            return low + ((high - low) / 2);
        }

        private sealed class Node(int center, Subscription[] byStart, Subscription[] byEnd, Node? left, Node? right)
        {
            public int Center { get; } = center;

            public Subscription[] ByStart { get; } = byStart;

            public Subscription[] ByEnd { get; } = byEnd;

            public Node? Left { get; } = left;

            public Node? Right { get; } = right;
        }
    }

    /// <summary>
    /// Collects subscriptions while the dependency graph is wired up, merging identical
    /// ranges: fifty thousand lookups against one table become a single entry.
    /// </summary>
    internal sealed class Builder
    {
        private readonly Dictionary<RangeKey, int> _payloadByRange = [];
        private readonly List<List<int>> _payloads = [];
        private readonly List<RangeKey> _ranges = [];

        public void Add(in RangeKey range, int entry)
        {
            if (_payloadByRange.TryGetValue(range, out int payload))
            {
                _payloads[payload].Add(entry);
                return;
            }

            payload = _payloads.Count;
            _payloadByRange[range] = payload;
            _payloads.Add([entry]);
            _ranges.Add(range);
        }

        public RangeSubscriptionIndex Build()
        {
            var bySheet = new Dictionary<int, SheetIndex>();
            var pending = new Dictionary<int, Dictionary<int, List<Subscription>>>();

            for (int payload = 0; payload < _ranges.Count; payload++)
            {
                RangeKey range = _ranges[payload];
                var subscription = new Subscription(
                    range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, payload);

                SheetIndex sheet = GetSheet(bySheet, range.SheetId);
                int width = range.LastColumn - range.FirstColumn + 1;
                if (width > MaxBucketColumns)
                {
                    sheet.Wide.Add(subscription);
                    continue;
                }

                Dictionary<int, List<Subscription>> columns = GetColumns(pending, range.SheetId);
                for (int column = range.FirstColumn; column <= range.LastColumn; column++)
                {
                    if (!columns.TryGetValue(column, out List<Subscription>? bucket))
                        columns[column] = bucket = [];
                    bucket.Add(subscription);
                }
            }

            foreach ((int sheetId, Dictionary<int, List<Subscription>> columns) in pending)
            {
                SheetIndex sheet = GetSheet(bySheet, sheetId);
                foreach ((int column, List<Subscription> bucket) in columns)
                    sheet.Columns[column] = IntervalTree.Build(bucket);
            }

            return new RangeSubscriptionIndex(bySheet, [.. _payloads]);
        }

        private static SheetIndex GetSheet(Dictionary<int, SheetIndex> bySheet, int sheetId)
        {
            if (!bySheet.TryGetValue(sheetId, out SheetIndex? sheet))
                bySheet[sheetId] = sheet = new SheetIndex();
            return sheet;
        }

        private static Dictionary<int, List<Subscription>> GetColumns(
            Dictionary<int, Dictionary<int, List<Subscription>>> pending, int sheetId)
        {
            if (!pending.TryGetValue(sheetId, out Dictionary<int, List<Subscription>>? columns))
                pending[sheetId] = columns = [];
            return columns;
        }
    }
}
