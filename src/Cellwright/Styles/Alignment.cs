namespace Cellwright;

/// <summary>Horizontal cell alignment.</summary>
public enum HorizontalAlignment : byte
{
    /// <summary>Type-dependent default: text left, numbers right.</summary>
    General = 0,

    /// <summary>Left aligned.</summary>
    Left,

    /// <summary>Centered.</summary>
    Center,

    /// <summary>Right aligned.</summary>
    Right,

    /// <summary>Justified.</summary>
    Justify,

    /// <summary>The content is repeated until it fills the cell.</summary>
    Fill,

    /// <summary>Centered across the neighbouring empty cells rather than within this one.</summary>
    CenterContinuous,

    /// <summary>Spread out so that every line reaches both edges, last line included.</summary>
    Distributed,
}

/// <summary>Vertical cell alignment.</summary>
public enum VerticalAlignment : byte
{
    /// <summary>Bottom aligned (Excel default).</summary>
    Bottom = 0,

    /// <summary>Top aligned.</summary>
    Top,

    /// <summary>Centered.</summary>
    Center,

    /// <summary>Lines spread to fill the height.</summary>
    Justify,

    /// <summary>Lines spread evenly, including above the first and below the last.</summary>
    Distributed,
}

/// <summary>Which way text runs inside a cell.</summary>
public enum ReadingOrder : byte
{
    /// <summary>Decided by the first strongly directional character.</summary>
    Context = 0,

    /// <summary>Left to right.</summary>
    LeftToRight,

    /// <summary>Right to left.</summary>
    RightToLeft,
}

/// <summary>
/// Immutable alignment settings of a cell style.
/// </summary>
public sealed record Alignment
{
    /// <summary>Excel defaults: general horizontal, bottom vertical, no wrap.</summary>
    public static Alignment Default { get; } = new();

    /// <summary>Horizontal alignment.</summary>
    public HorizontalAlignment Horizontal { get; init; }

    /// <summary>Vertical alignment.</summary>
    public VerticalAlignment Vertical { get; init; }

    /// <summary>Wrap long text onto multiple lines.</summary>
    public bool WrapText { get; init; }

    /// <summary>Indent level (0–250).</summary>
    public int Indent { get; init; }

    /// <summary>
    /// Rotation in degrees counter-clockwise, 0 to 90, or 91 to 180 for clockwise 1 to 90.
    /// The value 255 is Excel's marker for stacked letters reading downwards.
    /// </summary>
    public int TextRotation { get; init; }

    /// <summary>Shrink the font until the text fits the column rather than wrapping it.</summary>
    public bool ShrinkToFit { get; init; }

    /// <summary>Which way the text runs.</summary>
    public ReadingOrder ReadingOrder { get; init; }

    /// <summary>Excel's marker for text stacked letter under letter.</summary>
    public const int StackedRotation = 255;
}
