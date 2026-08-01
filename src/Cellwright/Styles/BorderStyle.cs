namespace Cellwright;

/// <summary>
/// Line style of one border edge, matching <c>ST_BorderStyle</c> (§18.18.3).
/// </summary>
public enum BorderLineStyle : byte
{
    /// <summary>No border.</summary>
    None = 0,

    /// <summary>Thin solid line.</summary>
    Thin,

    /// <summary>Medium solid line.</summary>
    Medium,

    /// <summary>Thick solid line.</summary>
    Thick,

    /// <summary>Dashed line.</summary>
    Dashed,

    /// <summary>Dotted line.</summary>
    Dotted,

    /// <summary>Double line.</summary>
    Double,

    /// <summary>Hairline.</summary>
    Hair,

    /// <summary>Medium dashed line.</summary>
    MediumDashed,

    /// <summary>Dash-dot line.</summary>
    DashDot,

    /// <summary>Medium dash-dot line.</summary>
    MediumDashDot,

    /// <summary>Dash-dot-dot line.</summary>
    DashDotDot,

    /// <summary>Medium dash-dot-dot line.</summary>
    MediumDashDotDot,

    /// <summary>Slanted dash-dot line.</summary>
    SlantDashDot,
}

/// <summary>
/// One border edge: a line style plus a color.
/// </summary>
/// <param name="Line">Line style; <see cref="BorderLineStyle.None"/> disables the edge.</param>
/// <param name="Color">Line color.</param>
public readonly record struct BorderEdge(BorderLineStyle Line, ExcelColor Color = default)
{
    /// <summary>A thin black edge — the most common grid border.</summary>
    public static BorderEdge Thin => new(BorderLineStyle.Thin);
}

/// <summary>
/// Immutable border settings of a cell style.
/// </summary>
public sealed record BorderStyle
{
    /// <summary>No borders.</summary>
    public static BorderStyle None { get; } = new();

    /// <summary>Left edge.</summary>
    public BorderEdge Left { get; init; }

    /// <summary>Right edge.</summary>
    public BorderEdge Right { get; init; }

    /// <summary>Top edge.</summary>
    public BorderEdge Top { get; init; }

    /// <summary>Bottom edge.</summary>
    public BorderEdge Bottom { get; init; }

    /// <summary>
    /// The diagonal line's style and color. Which corners it joins is said separately, by
    /// <see cref="DiagonalUp"/> and <see cref="DiagonalDown"/>; without either it draws nothing.
    /// </summary>
    public BorderEdge Diagonal { get; init; }

    /// <summary>Whether the diagonal runs from the bottom-left corner to the top-right one.</summary>
    public bool DiagonalUp { get; init; }

    /// <summary>Whether the diagonal runs from the top-left corner to the bottom-right one.</summary>
    public bool DiagonalDown { get; init; }

    /// <summary>Creates a border with the same edge on all four sides.</summary>
    public static BorderStyle All(BorderEdge edge) => new() { Left = edge, Right = edge, Top = edge, Bottom = edge };
}
