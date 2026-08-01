using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// What a cell shows and how: the text a number format renders, the colour that format may
/// override, and which way the text lines up in the cell.
/// </summary>
/// <param name="Text">The characters the cell shows, empty when it shows nothing.</param>
/// <param name="Alignment">Where the text sits across the cell, already resolved from General.</param>
/// <param name="Color">The colour to draw it in, which a format section may have overridden.</param>
/// <param name="IsNumeric">
/// Whether the value is one that has to fit: a number too wide for its column prints as hashes,
/// where text too wide spills or is clipped.
/// </param>
internal readonly record struct CellText(
    string Text, HorizontalAlignment Alignment, PdfColor Color, bool IsNumeric)
{
    /// <summary>Whether the cell shows nothing at all.</summary>
    public bool IsEmpty => Text.Length == 0;
}

/// <summary>Turns a cell's value and style into the text the page shows.</summary>
internal static class CellContent
{
    /// <summary>
    /// The colours a number format can name in brackets, which override the font's own
    /// (ISO/IEC 29500-1 §18.8.31). The eight names are the only ones a format may use;
    /// <c>[Color n]</c> indexes the legacy palette instead.
    /// </summary>
    private static readonly Dictionary<string, uint> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Black"] = 0x000000,
        ["White"] = 0xFFFFFF,
        ["Red"] = 0xFF0000,
        ["Green"] = 0x00B050,
        ["Blue"] = 0x0000FF,
        ["Yellow"] = 0xFFFF00,
        ["Magenta"] = 0xFF00FF,
        ["Cyan"] = 0x00FFFF,
    };

    /// <summary>Reads what a cell shows.</summary>
    /// <param name="context">The render in progress.</param>
    /// <param name="value">The cell's value.</param>
    /// <param name="style">The cell's style.</param>
    /// <param name="fallback">The colour the font would draw in.</param>
    public static CellText Of(PdfExportContext context, in CellValue value, CellStyle style, PdfColor fallback)
    {
        ArgumentNullException.ThrowIfNull(style);

        if (value.IsBlank)
            return new CellText("", HorizontalAlignment.General, fallback, IsNumeric: false);

        string text = NumberFormatRenderer.Format(
            value, style.NumberFormat, context.Culture, context.Source.DateSystem);

        bool numeric = value.IsNumber;
        PdfColor color = SectionColor(style.NumberFormat, value, numeric) is { } named
            ? PdfColor.FromRgb(named)
            : fallback;

        return new CellText(text, Resolve(style.Alignment.Horizontal, value), color, numeric);
    }

    /// <summary>
    /// Where the text sits when the style leaves it to the value: numbers and dates to the
    /// right, text to the left, and the values that are neither in the middle.
    /// </summary>
    private static HorizontalAlignment Resolve(HorizontalAlignment stated, in CellValue value)
    {
        if (stated != HorizontalAlignment.General)
            return stated;

        if (value.IsNumber)
            return HorizontalAlignment.Right;

        return value.IsBoolean || value.IsError ? HorizontalAlignment.Center : HorizontalAlignment.Left;
    }

    /// <summary>
    /// The colour the section that rendered this value names, if it names one. The sections are
    /// picked the same way the renderer picks them, so the colour belongs to the text on show.
    /// </summary>
    private static uint? SectionColor(string code, in CellValue value, bool numeric)
    {
        if (code.Length == 0 || code.IndexOf('[') < 0)
            return null;

        string[] sections = code.Split(';');
        int index = numeric ? NumericSection(sections.Length, value.GetNumber()) : Math.Min(3, sections.Length - 1);
        if (index < 0 || index >= sections.Length)
            return null;

        return Named(sections[index]);
    }

    private static int NumericSection(int count, double number) => count switch
    {
        1 => 0,
        2 => number < 0 ? 1 : 0,
        _ => number < 0 ? 1 : number == 0 ? 2 : 0,
    };

    /// <summary>The colour a section names in brackets, or null when it names none.</summary>
    private static uint? Named(string section)
    {
        int open = section.IndexOf('[');
        while (open >= 0)
        {
            int close = section.IndexOf(']', open);
            if (close < 0)
                return null;

            string inside = section[(open + 1)..close].Trim();
            if (NamedColors.TryGetValue(inside, out uint color))
                return color;

            if (inside.StartsWith("Color", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(inside.AsSpan(5).Trim(), out int slot))
            {
                ExcelColor indexed = ExcelColor.FromIndexed(slot);
                return indexed.IsAuto ? null : indexed.Argb & 0x00FFFFFFu;
            }

            open = section.IndexOf('[', close);
        }

        return null;
    }
}
