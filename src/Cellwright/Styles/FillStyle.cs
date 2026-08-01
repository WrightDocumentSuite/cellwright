namespace Cellwright;

/// <summary>
/// How a fill paints the cell, matching <c>ST_PatternType</c> (§18.18.55). Everything but
/// <see cref="Solid"/> mixes the foreground and background colors in a fixed pattern.
/// </summary>
public enum FillPattern : byte
{
    /// <summary>Nothing painted.</summary>
    None = 0,

    /// <summary>The foreground color, edge to edge.</summary>
    Solid,

    /// <summary>75% of the foreground color.</summary>
    DarkGray,

    /// <summary>50% of the foreground color.</summary>
    MediumGray,

    /// <summary>25% of the foreground color.</summary>
    LightGray,

    /// <summary>12.5% of the foreground color.</summary>
    Gray125,

    /// <summary>6.25% of the foreground color.</summary>
    Gray0625,

    /// <summary>Horizontal stripes.</summary>
    DarkHorizontal,

    /// <summary>Vertical stripes.</summary>
    DarkVertical,

    /// <summary>Diagonal stripes going down.</summary>
    DarkDown,

    /// <summary>Diagonal stripes going up.</summary>
    DarkUp,

    /// <summary>Crosshatch.</summary>
    DarkGrid,

    /// <summary>Diagonal crosshatch.</summary>
    DarkTrellis,

    /// <summary>Thin horizontal stripes.</summary>
    LightHorizontal,

    /// <summary>Thin vertical stripes.</summary>
    LightVertical,

    /// <summary>Thin diagonal stripes going down.</summary>
    LightDown,

    /// <summary>Thin diagonal stripes going up.</summary>
    LightUp,

    /// <summary>Thin crosshatch.</summary>
    LightGrid,

    /// <summary>Thin diagonal crosshatch.</summary>
    LightTrellis,
}

/// <summary>
/// Immutable fill of a cell style: a pattern drawn with a foreground and a background color,
/// or a <see cref="Gradient"/>. A solid fill uses only the foreground one, which is the case
/// nearly every file needs.
/// </summary>
public sealed record FillStyle
{
    /// <summary>No fill.</summary>
    public static FillStyle None { get; } = new();

    /// <summary>How the two colors are laid down. Ignored when <see cref="Gradient"/> is set.</summary>
    public FillPattern Pattern { get; init; }

    /// <summary>Foreground color: the paint of a solid fill, the lines of a pattern.</summary>
    public ExcelColor Color { get; init; } = ExcelColor.Auto;

    /// <summary>Background color, showing between the lines of a pattern.</summary>
    public ExcelColor BackgroundColor { get; init; } = ExcelColor.Auto;

    /// <summary>
    /// Gradient sweep (§18.8.24), or <see langword="null"/> for a patterned fill. A fill is one
    /// or the other, never both.
    /// </summary>
    public GradientFill? Gradient { get; init; }

    /// <summary>Creates a solid fill.</summary>
    public static FillStyle Solid(ExcelColor color) => new() { Pattern = FillPattern.Solid, Color = color };

    /// <summary>Creates a patterned fill.</summary>
    public static FillStyle Patterned(FillPattern pattern, ExcelColor foreground, ExcelColor background = default) =>
        new() { Pattern = pattern, Color = foreground, BackgroundColor = background };

    /// <summary>Creates a gradient fill.</summary>
    public static FillStyle Gradiented(GradientFill gradient) => new() { Gradient = gradient };

    /// <summary>Whether the fill is a plain block of one color.</summary>
    public bool IsSolid => Pattern == FillPattern.Solid && Gradient is null;

    /// <summary>Whether the fill paints anything at all.</summary>
    public bool IsEmpty => Pattern == FillPattern.None && Gradient is null;
}
