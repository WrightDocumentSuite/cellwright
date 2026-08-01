namespace Cellwright.Pdf.Layout;

/// <summary>
/// Where every printed line of one page sits along one axis: which row or column of the sheet it
/// is, and where its near edge falls, with the repeated titles already at the front.
/// </summary>
/// <remarks>
/// Positions are relative to the top-left of the grid on the page, before the page's own offset
/// and scale, so the same slots can be drawn at any size without being measured again.
/// </remarks>
internal sealed class PageSlots
{
    private (int Least, int Most)? _range;

    /// <summary>The rows or columns of the sheet, in the order this page draws them.</summary>
    public required int[] Lines { get; init; }

    /// <summary>Where each line sits among the sheet's printed ones.</summary>
    public required int[] Indices { get; init; }

    /// <summary>The near edge of each line, in points from the grid's own origin.</summary>
    public required double[] Near { get; init; }

    /// <summary>How wide or tall each line is.</summary>
    public required double[] Size { get; init; }

    /// <summary>
    /// How many of the leading slots are repeated titles rather than the band this page is of.
    /// </summary>
    public required int Titles { get; init; }

    /// <summary>Whether the band starts at the beginning of the sheet's printed lines.</summary>
    public required bool First { get; init; }

    /// <summary>Whether it runs to the end of them.</summary>
    public required bool Last { get; init; }

    /// <summary>How many lines this page draws along the axis.</summary>
    public int Count => Lines.Length;

    /// <summary>Whether any line this page draws falls within a run of them.</summary>
    /// <param name="first">The first line of the run, 1-based.</param>
    /// <param name="last">The last line of the run.</param>
    /// <remarks>
    /// The repeated titles come before the band, so the lines are not in order and the ends of
    /// the array are not the ends of the range. The range is worked out once and kept: a sheet
    /// asks this of every merge it has, for every page it comes to.
    /// </remarks>
    public bool Meets(int first, int last)
    {
        if (Count == 0)
            return false;

        (int least, int most) = _range ??= Range();
        return last >= least && first <= most;
    }

    private (int Least, int Most) Range()
    {
        int least = int.MaxValue;
        int most = int.MinValue;

        foreach (int line in Lines)
        {
            least = Math.Min(least, line);
            most = Math.Max(most, line);
        }

        return (least, most);
    }

    /// <summary>The far edge of a line.</summary>
    /// <param name="slot">The position among this page's lines.</param>
    public double Far(int slot) => Near[slot] + Size[slot];

    /// <summary>How far the whole axis reaches.</summary>
    public double Extent => Count == 0 ? 0 : Far(Count - 1) - Near[0];

    /// <summary>Which slot holds a row or column of the sheet, or -1 when this page has none.</summary>
    /// <param name="line">The row or column number, 1-based.</param>
    public int SlotOf(int line) => Array.IndexOf(Lines, line);
}
