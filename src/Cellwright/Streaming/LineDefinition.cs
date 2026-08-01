namespace Cellwright.Streaming;

/// <summary>
/// One <c>row</c> element's properties as read from the file. The style index is the raw
/// styles-part index; the loader maps it into the workbook's pool the same way it maps cells.
/// </summary>
/// <param name="Row">1-based row number.</param>
/// <param name="Height">Explicit height in points, or <see langword="null"/> for the sheet default.</param>
/// <param name="Hidden">Whether the row is hidden.</param>
/// <param name="OutlineLevel">Outline (grouping) level.</param>
/// <param name="Collapsed">Whether the group ending here is shown collapsed.</param>
/// <param name="RawStyleIndex">Style index into the source styles part, or 0 for none.</param>
/// <param name="ThickTop">Whether the top edge is drawn thick.</param>
/// <param name="ThickBottom">Whether the bottom edge is drawn thick.</param>
internal readonly record struct RowDefinition(
    int Row,
    double? Height,
    bool Hidden,
    int OutlineLevel,
    bool Collapsed,
    int RawStyleIndex,
    bool ThickTop,
    bool ThickBottom)
{
    /// <summary>Whether the element said anything worth keeping.</summary>
    public bool IsEmpty =>
        Height is null && !Hidden && OutlineLevel == 0 && !Collapsed &&
        RawStyleIndex == 0 && !ThickTop && !ThickBottom;
}

/// <summary>One <c>col</c> element's properties as read from the file.</summary>
/// <param name="Column">1-based column number.</param>
/// <param name="Width">Explicit width in characters, or <see langword="null"/> for the sheet default.</param>
/// <param name="Hidden">Whether the column is hidden.</param>
/// <param name="OutlineLevel">Outline (grouping) level.</param>
/// <param name="Collapsed">Whether the group ending here is shown collapsed.</param>
/// <param name="RawStyleIndex">Style index into the source styles part, or 0 for none.</param>
/// <param name="BestFit">Whether Excel sized the column to its content.</param>
internal readonly record struct ColumnDefinition(
    int Column,
    double? Width,
    bool Hidden,
    int OutlineLevel,
    bool Collapsed,
    int RawStyleIndex,
    bool BestFit)
{
    /// <inheritdoc cref="RowDefinition.IsEmpty"/>
    public bool IsEmpty =>
        Width is null && !Hidden && OutlineLevel == 0 && !Collapsed && RawStyleIndex == 0 && !BestFit;
}
