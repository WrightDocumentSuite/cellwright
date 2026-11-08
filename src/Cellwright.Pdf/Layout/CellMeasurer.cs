using Cellwright.Styles;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// How much room one cell's text asks for, in the points a page is measured in.
/// </summary>
/// <remarks>
/// Pagination and auto-fitting both need this answer and they must not answer it differently: a
/// column sized to hold its text and then printed a hair narrower cuts the text it was sized for.
/// So both go through here, with the same font metrics, the same number formatting and the same
/// padding the painter later leaves around the glyphs.
/// </remarks>
internal sealed class CellMeasurer
{
    /// <summary>Room left between a cell's edge and its text, each side, as the painter leaves it.</summary>
    public const double Padding = 2;

    /// <summary>How far one step of indent moves the text in.</summary>
    public const double IndentStep = 9;

    /// <summary>Room left above and below a line, which is what keeps rows off each other's glyphs.</summary>
    private const double RowPadding = 2;

    private readonly PdfExportContext _context;
    private readonly Worksheet _sheet;

    private readonly ConditionalFormatEvaluator? _conditional;

    internal CellMeasurer(
        PdfExportContext context,
        Worksheet sheet,
        CellAppearances appearances,
        ConditionalFormatEvaluator? conditional = null)
    {
        _context = context;
        _sheet = sheet;
        _conditional = conditional;
        Appearances = appearances;
    }

    /// <summary>The appearances the sheet's cells are drawn in.</summary>
    public CellAppearances Appearances { get; }

    /// <summary>
    /// The style a cell is measured in: the one it states, with whatever conditional formatting
    /// matched laid over it.
    /// </summary>
    /// <param name="row">Row number, 1-based.</param>
    /// <param name="column">Column number, 1-based.</param>
    /// <remarks>
    /// A rule can change the number format and the font, and both change how much room the text
    /// takes. Measuring the stated style and drawing the conditioned one would size a row to text
    /// nobody prints.
    /// </remarks>
    public CellStyle StyleOf(int row, int column)
    {
        CellStyle stated = _sheet.Workbook.GetStyle(new StyleId(_sheet.StyleIndexes.Get(row, column)));
        return _conditional is { HasRules: true } evaluator
            ? evaluator.Overlay(stated, evaluator.Of(new CellRef(row, column)))
            : stated;
    }

    /// <summary>The text a cell shows, formatted the way the page will print it.</summary>
    /// <param name="value">The cell's value.</param>
    /// <param name="style">The style it wears.</param>
    public string TextOf(in CellValue value, CellStyle style) =>
        NumberFormatRenderer.Format(value, style.NumberFormat, _context.Culture, _context.Source.DateSystem);

    /// <summary>
    /// How tall a cell needs its row to be, given the room its column gives the text.
    /// </summary>
    /// <param name="value">The cell's value.</param>
    /// <param name="row">Row number, 1-based.</param>
    /// <param name="column">Column number, 1-based.</param>
    /// <param name="columnWidth">How wide the column is, in points.</param>
    /// <remarks>
    /// Only wrapped, unturned text can need more than a line: everything else is drawn on one,
    /// whether or not it fits, and a cell that overflows spills sideways rather than downwards.
    /// </remarks>
    public double HeightNeeded(in CellValue value, int row, int column, double columnWidth)
    {
        CellStyle style = StyleOf(row, column);
        TextAppearance appearance = Appearances.Of(style.Font);

        if (style.Alignment.TextRotation != 0)
        {
            int degrees = style.Alignment.TextRotation;
            if (degrees is (>= 1 and < 90) or (> 90 and < 180))
            {
                string rotatedText = TextOf(value, style);
                if (rotatedText.Length > 0)
                {
                    double angleDeg = degrees <= 90 ? degrees : degrees - 90;
                    double rad = angleDeg * Math.PI / 180.0;
                    double w = appearance.Measure(rotatedText);
                    double h = appearance.LineHeight;
                    double projectedHeight = (w * Math.Sin(rad)) + (h * Math.Cos(rad));
                    return projectedHeight + RowPadding;
                }
            }

            return appearance.LineHeight + RowPadding;
        }

        string text = TextOf(value, style);
        if (text.Length == 0)
            return appearance.LineHeight + RowPadding;

        double room = Room(columnWidth, style.Alignment);
        if (_sheet.Cell(row, column).RichText is { IsFormatted: true } rich &&
            string.Equals(text, rich.Text, StringComparison.Ordinal))
        {
            List<(double Width, double LineHeight)> richLines =
                CellPlacer.MeasureRichLines(rich, style.Font, Appearances, style.Alignment, room);
            if (!style.Alignment.WrapText)
            {
                double maxLineHeight = appearance.LineHeight;
                foreach ((_, double lineHeight) in richLines)
                    maxLineHeight = Math.Max(maxLineHeight, lineHeight);
                return maxLineHeight + RowPadding;
            }

            double totalHeight = 0;
            foreach ((_, double lineHeight) in richLines)
                totalHeight += lineHeight;
            return totalHeight + RowPadding;
        }

        // Every cell wants room for one line of its own font, wrapped or not.
        if (!style.Alignment.WrapText)
            return appearance.LineHeight + RowPadding;

        int lines = CellPlacer.CountWrappedLines(text, appearance, room, style.Alignment);
        return (lines * appearance.LineHeight) + RowPadding;
    }

    /// <summary>
    /// How wide a cell needs its column to be for its text to be drawn whole on one line,
    /// padding and indent included.
    /// </summary>
    /// <param name="value">The cell's value.</param>
    /// <param name="row">Row number, 1-based.</param>
    /// <param name="column">Column number, 1-based.</param>
    /// <remarks>
    /// Text turned on its side asks for the width of a line rather than the width of the text —
    /// what runs across the column is then the height of the glyphs, and the length of the string
    /// runs down the row instead.
    /// </remarks>
    public double WidthNeeded(in CellValue value, int row, int column)
    {
        CellStyle style = StyleOf(row, column);
        TextAppearance appearance = Appearances.Of(style.Font);

        string text = TextOf(value, style);
        if (text.Length == 0)
            return 0;

        if (style.Alignment.TextRotation == Alignment.StackedRotation)
        {
            // Stacked text puts one upright letter under the next, so what runs across the
            // column is the widest single letter rather than the string.
            double widestLetter = 0;
            foreach (char letter in text)
                widestLetter = Math.Max(widestLetter, appearance.Measure(letter.ToString()));

            return widestLetter + (2 * Padding);
        }

        if (style.Alignment.TextRotation != 0)
        {
            int degrees = style.Alignment.TextRotation;
            if (degrees is (>= 1 and < 90) or (> 90 and < 180))
            {
                double angleDeg = degrees <= 90 ? degrees : degrees - 90;
                double rad = angleDeg * Math.PI / 180.0;
                double w = appearance.Measure(text);
                double h = appearance.LineHeight;
                double projectedWidth = (w * Math.Cos(rad)) + (h * Math.Sin(rad));
                return projectedWidth + (2 * Padding);
            }

            return appearance.LineHeight + (2 * Padding);
        }

        if (_sheet.Cell(row, column).RichText is { IsFormatted: true } rich &&
            string.Equals(text, rich.Text, StringComparison.Ordinal))
        {
            List<(double Width, double LineHeight)> richLines =
                CellPlacer.MeasureRichLines(
                    rich, style.Font, Appearances, style.Alignment with { WrapText = false }, double.MaxValue);
            double widestRich = 0;
            foreach ((double width, _) in richLines)
                widestRich = Math.Max(widestRich, width);
            return widestRich + (2 * Padding) + (style.Alignment.Indent * IndentStep);
        }

        // A cell holding several lines is as wide as its widest one. Splitting is worth avoiding
        // in the case that is almost every cell: one line, already a string, measured as it is.
        double widest;
        if (text.Contains('\n'))
        {
            widest = 0;
            foreach (ReadOnlySpan<char> line in text.AsSpan().EnumerateLines())
                widest = Math.Max(widest, appearance.Measure(line.ToString()));
        }
        else
        {
            widest = appearance.Measure(text);
        }

        return widest + (2 * Padding) + (style.Alignment.Indent * IndentStep);
    }

    /// <summary>How much of a column's width the text may actually use.</summary>
    /// <param name="columnWidth">How wide the column is, in points.</param>
    /// <param name="alignment">The cell's alignment, whose indent eats into the room.</param>
    public static double Room(double columnWidth, Alignment alignment) =>
        Math.Max(1, columnWidth - (2 * Padding) - (alignment.Indent * IndentStep));
}
