namespace Cellwright;

/// <summary>
/// Everything a worksheet stores about one row besides its cells: an explicit height,
/// visibility, the outline level that puts it under a grouping bracket, and the style its
/// empty cells inherit.
/// </summary>
public readonly record struct RowProperties
{
    /// <summary>Height in points, or <see langword="null"/> to use the sheet default.</summary>
    public double? Height { get; init; }

    /// <summary>Whether the row is hidden.</summary>
    public bool Hidden { get; init; }

    /// <summary>Outline (grouping) level, 0 through 7.</summary>
    public int OutlineLevel { get; init; }

    /// <summary>Whether the outline group ending at this row is shown collapsed.</summary>
    public bool Collapsed { get; init; }

    /// <summary>
    /// Style the row's cells inherit where they carry none of their own. This is what makes a
    /// whole formatted row cost one entry rather than sixteen thousand.
    /// </summary>
    public StyleId Style { get; init; }

    /// <summary>Whether the row's top edge is drawn thick, which an outline border above it needs.</summary>
    public bool ThickTop { get; init; }

    /// <summary>Whether the row's bottom edge is drawn thick.</summary>
    public bool ThickBottom { get; init; }

    /// <summary>Whether the row differs from the sheet default in any way.</summary>
    public bool IsDefault =>
        Height is null && !Hidden && OutlineLevel == 0 && !Collapsed &&
        Style == StyleId.Default && !ThickTop && !ThickBottom;
}

/// <summary>
/// Everything a worksheet stores about one column besides its cells: an explicit width,
/// visibility, the outline level that puts it under a grouping bracket, and the style its
/// cells inherit.
/// </summary>
public readonly record struct ColumnProperties
{
    /// <summary>Width in characters, or <see langword="null"/> to use the sheet default.</summary>
    public double? Width { get; init; }

    /// <summary>Whether the column is hidden.</summary>
    public bool Hidden { get; init; }

    /// <summary>Outline (grouping) level, 0 through 7.</summary>
    public int OutlineLevel { get; init; }

    /// <summary>Whether the outline group ending at this column is shown collapsed.</summary>
    public bool Collapsed { get; init; }

    /// <summary>Style the column's cells inherit where they carry none of their own.</summary>
    public StyleId Style { get; init; }

    /// <summary>Whether Excel sized the column to its content rather than being told a width.</summary>
    public bool BestFit { get; init; }

    /// <summary>Whether the column differs from the sheet default in any way.</summary>
    public bool IsDefault =>
        Width is null && !Hidden && OutlineLevel == 0 && !Collapsed && Style == StyleId.Default && !BestFit;
}
