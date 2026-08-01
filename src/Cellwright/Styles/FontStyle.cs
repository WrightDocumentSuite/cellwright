namespace Cellwright;

/// <summary>How a font is underlined (<c>ST_UnderlineValues</c>, §18.18.85).</summary>
public enum UnderlineStyle : byte
{
    /// <summary>Not underlined.</summary>
    None = 0,

    /// <summary>One line under the text.</summary>
    Single,

    /// <summary>Two lines under the text.</summary>
    Double,

    /// <summary>One line running the full cell width, as accountants rule a column.</summary>
    SingleAccounting,

    /// <summary>Two lines running the full cell width.</summary>
    DoubleAccounting,
}

/// <summary>Whether the text sits on the baseline or beside it.</summary>
public enum FontVerticalAlign : byte
{
    /// <summary>On the baseline.</summary>
    Baseline = 0,

    /// <summary>Raised and smaller.</summary>
    Superscript,

    /// <summary>Lowered and smaller.</summary>
    Subscript,
}

/// <summary>Which slot of the document theme a font follows.</summary>
public enum FontScheme : byte
{
    /// <summary>None: the font is named outright and does not change with the theme.</summary>
    None = 0,

    /// <summary>The theme's heading font.</summary>
    Major,

    /// <summary>The theme's body font.</summary>
    Minor,
}

/// <summary>
/// Immutable font settings of a cell style.
/// </summary>
public sealed record FontStyle
{
    /// <summary>The workbook default: Calibri 11.</summary>
    public static FontStyle Default { get; } = new();

    /// <summary>Font family name.</summary>
    public string Name { get; init; } = "Calibri";

    /// <summary>Size in points.</summary>
    public double Size { get; init; } = 11;

    /// <summary>Bold flag.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic flag.</summary>
    public bool Italic { get; init; }

    /// <summary>How the text is underlined.</summary>
    public UnderlineStyle Underline { get; init; }

    /// <summary>Strikethrough flag.</summary>
    public bool Strikethrough { get; init; }

    /// <summary>Whether the text sits on the baseline or beside it.</summary>
    public FontVerticalAlign VerticalAlign { get; init; }

    /// <summary>Text color.</summary>
    public ExcelColor Color { get; init; } = ExcelColor.Auto;

    /// <summary>
    /// Font family class (1 roman, 2 swiss, 3 modern, 4 script, 5 decorative), or 0 when the
    /// file said nothing. It only guides substitution when the named font is missing.
    /// </summary>
    public int Family { get; init; }

    /// <summary>Windows character set number, or -1 when the file said nothing.</summary>
    public int Charset { get; init; } = -1;

    /// <summary>
    /// Theme slot the font follows. A font that names one changes with the document theme, so
    /// dropping it would freeze the workbook's typography at whatever it looked like on load.
    /// </summary>
    public FontScheme Scheme { get; init; }
}
