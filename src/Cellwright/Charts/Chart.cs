using System.Collections;

namespace Cellwright;

/// <summary>What a chart plots its series as, named after the chartSpace elements (§21.2.2).</summary>
public enum ChartKind
{
    /// <summary>A kind this model does not name; the raw markup still travels with the chart.</summary>
    Unknown,

    /// <summary>Vertical bars — what Excel calls a column chart.</summary>
    Column,

    /// <summary>Horizontal bars.</summary>
    Bar,

    /// <summary>Lines through the points.</summary>
    Line,

    /// <summary>One ring of shares.</summary>
    Pie,

    /// <summary>A pie with a hole.</summary>
    Doughnut,

    /// <summary>Filled area under the line.</summary>
    Area,

    /// <summary>Points by two values, optionally joined.</summary>
    Scatter,

    /// <summary>Spokes from the centre.</summary>
    Radar,
}

/// <summary>
/// One series of a chart: where its name, categories and values come from, and what the file
/// last computed for them.
/// </summary>
/// <remarks>
/// A series speaks in references — <c>Лист1!$B$2:$B$10</c> — and carries a cache of the values
/// those references held when the chart was last drawn. Both are kept: the references are what
/// the series means, the caches are what it showed, and a reader without the workbook at hand
/// (or with a reference into a sheet since renamed) still has something honest to work with.
/// </remarks>
public sealed record ChartSeries
{
    /// <summary>The series name, as cached text, or <see langword="null"/>.</summary>
    public string? Name { get; init; }

    /// <summary>Reference the name comes from, or <see langword="null"/> when it was typed in.</summary>
    public string? NameReference { get; init; }

    /// <summary>Reference the category labels come from, or <see langword="null"/>.</summary>
    public string? CategoriesReference { get; init; }

    /// <summary>Reference the values come from, or <see langword="null"/>.</summary>
    public string? ValuesReference { get; init; }

    /// <summary>The category labels the file cached, in point order. Empty when it cached none.</summary>
    public IReadOnlyList<string> Categories { get; init; } = [];

    /// <summary>The values the file cached, in point order; a null is a point with no value.</summary>
    public IReadOnlyList<double?> Values { get; init; } = [];
}

/// <summary>
/// A chart embedded on a worksheet.
/// </summary>
/// <remarks>
/// The raw <c>chartSpace</c> markup is the source of truth: it holds everything §21.2 can say,
/// which is far more than this model names, and it is what a save writes back as long as the
/// chart is untouched. The typed properties are a parsed view of it, built on first use —
/// reading a workbook costs nothing for charts nobody looks at.
/// </remarks>
public sealed class Chart
{
    private ChartView? _view;

    /// <summary>The chartSpace part, byte for byte as it came from the file.</summary>
    internal ReadOnlyMemory<byte> RawXml { get; init; }

    /// <summary>Part path the chart came from, so an untouched one writes back where it was.</summary>
    internal string? SourcePartPath { get; init; }

    /// <summary>Where the chart sits on its sheet.</summary>
    public required PictureAnchor Anchor { get; init; }

    /// <summary>The frame's name, e.g. <c>Chart 1</c>, or <see langword="null"/>.</summary>
    public string? Name { get; init; }

    /// <summary>What the chart plots its first group of series as.</summary>
    public ChartKind Kind => View.Kind;

    /// <summary>The chart title, when it is written-in text; <see langword="null"/> otherwise.</summary>
    public string? Title => View.Title;

    /// <summary>The series, in plot order.</summary>
    public IReadOnlyList<ChartSeries> Series => View.Series;

    private ChartView View => _view ??= ChartPart.Parse(RawXml);

    /// <inheritdoc/>
    public override string ToString() => $"{Kind} chart{(Name is null ? "" : $" '{Name}'")}";
}

/// <summary>The parsed view of a chartSpace, built once per chart on first use.</summary>
internal sealed record ChartView(ChartKind Kind, string? Title, IReadOnlyList<ChartSeries> Series)
{
    public static ChartView Empty { get; } = new(ChartKind.Unknown, null, []);
}

/// <summary>The charts embedded on one worksheet, in the order their anchors appear.</summary>
public sealed class ChartCollection : IReadOnlyList<Chart>
{
    private readonly List<Chart> _charts = [];

    /// <inheritdoc/>
    public int Count => _charts.Count;

    /// <inheritdoc/>
    public Chart this[int index] => _charts[index];

    /// <summary>
    /// Whether the model, rather than the loaded file, owns the sheet's drawing. Adding a chart
    /// means the drawing is written afresh, charts and pictures both.
    /// </summary>
    internal bool IsModified { get; private set; }

    /// <summary>Adds a chart read from a file.</summary>
    internal void AddFromFile(Chart chart) => _charts.Add(chart);

    /// <summary>Adds a chart the model built.</summary>
    internal Chart Add(Chart chart)
    {
        _charts.Add(chart);
        IsModified = true;
        return chart;
    }

    /// <summary>Removes a chart. Returns whether anything was removed.</summary>
    public bool Remove(Chart chart)
    {
        bool removed = _charts.Remove(chart);
        IsModified |= removed;
        return removed;
    }

    /// <inheritdoc/>
    public IEnumerator<Chart> GetEnumerator() => _charts.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
