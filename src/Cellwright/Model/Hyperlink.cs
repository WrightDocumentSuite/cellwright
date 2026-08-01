using System.Collections;

namespace Cellwright;

/// <summary>
/// A clickable region of a worksheet. Excel splits these in two: an external link stores a
/// URL in the sheet's relationships, while an internal one stores a cell reference in the
/// <c>location</c> attribute. Both shapes live here; <see cref="IsExternal"/> tells them apart.
/// </summary>
public sealed class Hyperlink
{
    private Hyperlink(RangeRef range, string? target, string? location, string? tooltip, string? display)
    {
        Range = range;
        Target = target;
        Location = location;
        Tooltip = tooltip;
        Display = display;
    }

    /// <summary>Cells the link covers.</summary>
    public RangeRef Range { get; internal set; }

    /// <summary>External destination (a URL or a file path), or <see langword="null"/> for internal links.</summary>
    public string? Target { get; init; }

    /// <summary>
    /// Location inside the destination: a cell reference or defined name such as
    /// <c>Sheet2!A1</c> for internal links, or a fragment of the external target.
    /// </summary>
    public string? Location { get; internal set; }

    /// <summary>Text shown while hovering, or <see langword="null"/>.</summary>
    /// <remarks>
    /// Settable from inside the library because the legacy format states it in a record of its
    /// own that follows the link, so it only becomes known after the link has been built.
    /// </remarks>
    public string? Tooltip { get; internal set; }

    /// <summary>Display text recorded alongside the link; the cell value is what Excel actually shows.</summary>
    public string? Display { get; init; }

    /// <summary>Whether the link leaves the workbook.</summary>
    public bool IsExternal => Target is not null;

    /// <summary>Creates a link to a URL or file path.</summary>
    public static Hyperlink ToUrl(RangeRef range, string url, string? tooltip = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(url);
        return new Hyperlink(range, url, location: null, tooltip, display: null);
    }

    /// <summary>Creates a link to a place in this workbook, e.g. <c>Summary!A1</c> or a defined name.</summary>
    public static Hyperlink ToLocation(RangeRef range, string location, string? tooltip = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(location);
        return new Hyperlink(range, target: null, location, tooltip, display: null);
    }

    internal static Hyperlink Load(RangeRef range, string? target, string? location, string? tooltip, string? display) =>
        new(range, target, location, tooltip, display);

    /// <summary>A separate link with the same destination, for a sheet that was duplicated.</summary>
    internal Hyperlink Clone() => new(Range, Target, Location, Tooltip, Display);

    /// <inheritdoc/>
    public override string ToString() => $"{Range} -> {Target ?? Location}";
}

/// <summary>
/// The hyperlinks of one worksheet. Adding a link over cells that already carry one replaces
/// it, matching what Excel does when you re-link a range.
/// </summary>
public sealed class HyperlinkCollection : IReadOnlyList<Hyperlink>
{
    private readonly List<Hyperlink> _links = [];

    /// <inheritdoc/>
    public int Count => _links.Count;

    /// <inheritdoc/>
    public Hyperlink this[int index] => _links[index];

    /// <summary>Adds a link, replacing any link on exactly the same range.</summary>
    public Hyperlink Add(Hyperlink link)
    {
        ArgumentNullException.ThrowIfNull(link);
        _links.RemoveAll(existing => existing.Range == link.Range);
        _links.Add(link);
        return link;
    }

    /// <summary>Adds a link to a URL over a range.</summary>
    public Hyperlink Add(RangeRef range, string url, string? tooltip = null) => Add(Hyperlink.ToUrl(range, url, tooltip));

    /// <summary>The link covering a cell, or <see langword="null"/>.</summary>
    public Hyperlink? Find(CellRef cell)
    {
        foreach (Hyperlink link in _links)
        {
            if (link.Range.Contains(cell))
                return link;
        }

        return null;
    }

    /// <summary>Removes every link intersecting a range. Returns how many were removed.</summary>
    public int Remove(RangeRef range) => _links.RemoveAll(link => link.Range.Intersects(range));

    /// <summary>
    /// Moves every link to the range the mapping gives it, dropping the ones it maps to
    /// nothing. Used when rows or columns are inserted or deleted underneath them.
    /// </summary>
    internal void Reanchor(Func<RangeRef, RangeRef?> map)
    {
        for (int i = _links.Count - 1; i >= 0; i--)
        {
            if (map(_links[i].Range) is { } moved)
                _links[i].Range = moved;
            else
                _links.RemoveAt(i);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<Hyperlink> GetEnumerator() => _links.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
