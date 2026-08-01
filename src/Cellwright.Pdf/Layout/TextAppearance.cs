using Cellwright.Styles;
using Inkwright;
using Inkwright.Fonts;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// How a piece of text is drawn, with the metrics that decide how much room it takes.
/// </summary>
/// <remarks>
/// A workbook has a handful of distinct appearances and any number of cells wearing them, so an
/// appearance is built once per distinct font style and shared. Measuring goes through here too,
/// which keeps the font's metrics in one place.
/// </remarks>
internal sealed record TextAppearance
{
    /// <summary>The face to draw with.</summary>
    public required PdfFont Font { get; init; }

    /// <summary>The size in points.</summary>
    public required double Size { get; init; }

    /// <summary>The colour of the glyphs.</summary>
    public required PdfColor Color { get; init; }

    /// <summary>How far the glyphs reach above the baseline.</summary>
    public required double Ascent { get; init; }

    /// <summary>How far they reach below it.</summary>
    public required double Descent { get; init; }

    /// <summary>How tall one line of this text is.</summary>
    public double LineHeight => Ascent + Descent;

    /// <summary>Whether a rule is drawn under the text, and how.</summary>
    public UnderlineStyle Underline { get; init; }

    /// <summary>Whether a rule is drawn through it.</summary>
    public bool Strikethrough { get; init; }

    /// <summary>How far the text is raised or lowered from the baseline.</summary>
    public double Rise { get; init; }

    /// <summary>How wide a stretch of text is at this size.</summary>
    /// <param name="text">The text to measure.</param>
    public double Measure(string text) => Font.Measure(text, Size);

    /// <summary>How wide a digit is, which is what a column's width counts in.</summary>
    public double DigitWidth => Measure("0");
}
