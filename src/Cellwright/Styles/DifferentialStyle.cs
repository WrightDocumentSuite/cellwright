namespace Cellwright;

/// <summary>
/// A partial style — what Excel calls a differential format (<c>dxf</c>). Only the properties
/// that are set take effect; everything else keeps the cell's own formatting. Conditional
/// formatting paints with these, which is why a rule can turn a cell red without also
/// resetting its font and borders.
/// </summary>
/// <example>
/// <code>
/// var overdue = new DifferentialStyle { Bold = true, FontColor = ExcelColor.FromRgb(0x9C, 0, 0x06) };
/// </code>
/// </example>
public sealed record DifferentialStyle
{
    /// <summary>A format that changes nothing.</summary>
    public static DifferentialStyle Empty { get; } = new();

    /// <summary>Bold override, or <see langword="null"/> to leave the cell's weight alone.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic override.</summary>
    public bool? Italic { get; init; }

    /// <summary>Single underline override.</summary>
    public bool? Underline { get; init; }

    /// <summary>Strikethrough override.</summary>
    public bool? Strikethrough { get; init; }

    /// <summary>Text color; <see cref="ExcelColor.Auto"/> leaves the cell's color alone.</summary>
    public ExcelColor FontColor { get; init; } = ExcelColor.Auto;

    /// <summary>Solid background color; <see cref="ExcelColor.Auto"/> leaves the cell's fill alone.</summary>
    public ExcelColor FillColor { get; init; } = ExcelColor.Auto;

    /// <summary>Number format code, e.g. <c>0.00%</c>, or <see langword="null"/> to keep the cell's.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>Border override, or <see langword="null"/> to keep the cell's borders.</summary>
    public BorderStyle? Border { get; init; }

    /// <summary>Whether the format would change anything at all.</summary>
    public bool IsEmpty =>
        Bold is null && Italic is null && Underline is null && Strikethrough is null &&
        FontColor.IsAuto && FillColor.IsAuto && NumberFormat is null && Border is null;

    /// <summary>Red text on a light red fill — the alert Excel offers first.</summary>
    public static DifferentialStyle Bad { get; } = new()
    {
        FontColor = ExcelColor.FromRgb(0x9C, 0x00, 0x06),
        FillColor = ExcelColor.FromRgb(0xFF, 0xC7, 0xCE),
    };

    /// <summary>Dark green text on a light green fill.</summary>
    public static DifferentialStyle Good { get; } = new()
    {
        FontColor = ExcelColor.FromRgb(0x00, 0x61, 0x00),
        FillColor = ExcelColor.FromRgb(0xC6, 0xEF, 0xCE),
    };

    /// <summary>Dark yellow text on a light yellow fill.</summary>
    public static DifferentialStyle Neutral { get; } = new()
    {
        FontColor = ExcelColor.FromRgb(0x9C, 0x65, 0x00),
        FillColor = ExcelColor.FromRgb(0xFF, 0xEB, 0x9C),
    };

    /// <summary>
    /// This format laid over a cell's own style: the properties it sets replace the cell's, and
    /// everything it leaves alone stays as the cell had it.
    /// </summary>
    /// <param name="baseStyle">The style the cell wears without the format.</param>
    /// <remarks>
    /// This is what a conditional format actually does to a cell, and the reason a rule can turn
    /// a cell red without also resetting its font, its alignment and its borders. A fill is
    /// replaced outright rather than blended: the format names a solid colour, so a cell whose
    /// own fill was a pattern or a gradient loses it while the rule holds, which is what Excel
    /// shows.
    /// </remarks>
    public CellStyle ApplyTo(CellStyle baseStyle)
    {
        ArgumentNullException.ThrowIfNull(baseStyle);
        if (IsEmpty)
            return baseStyle;

        FontStyle font = baseStyle.Font;
        if (Bold is bool bold)
            font = font with { Bold = bold };
        if (Italic is bool italic)
            font = font with { Italic = italic };
        if (Strikethrough is bool strikethrough)
            font = font with { Strikethrough = strikethrough };
        if (Underline is bool underline)
            font = font with { Underline = underline ? UnderlineStyle.Single : UnderlineStyle.None };
        if (!FontColor.IsAuto)
            font = font with { Color = FontColor };

        return baseStyle with
        {
            Font = font,
            Fill = FillColor.IsAuto ? baseStyle.Fill : FillStyle.Solid(FillColor),
            Border = Border ?? baseStyle.Border,
            NumberFormat = NumberFormat ?? baseStyle.NumberFormat,
        };
    }
}
