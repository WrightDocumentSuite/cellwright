using Cellwright.Styles;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// The two ways a cell's text stops being one straight line: wrapped onto several, or turned on
/// its side.
/// </summary>
internal sealed partial class CellPlacer
{
    /// <summary>
    /// How many lines a cell's text wraps onto, which is what a row with no stated height has
    /// to grow to hold.
    /// </summary>
    /// <param name="text">The text the cell shows.</param>
    /// <param name="appearance">How it is drawn.</param>
    /// <param name="room">How wide the lines may be.</param>
    /// <param name="alignment">The cell's alignment.</param>
    public static int CountWrappedLines(string text, TextAppearance appearance, double room, Alignment alignment) =>
        Break(text, appearance, room, alignment).Count;

    /// <summary>Draws text that wraps, one line per line box, top to bottom.</summary>
    private void Wrapped(
        CellText text, Alignment alignment, TextAppearance appearance, CellBox cell, double room, TagRef? tag)
    {
        List<string> lines = Break(text.Text, appearance, room, alignment);
        double total = lines.Count * appearance.LineHeight;
        double top = Down(alignment.Vertical, cell.Y, cell.Height, total);

        foreach (string line in lines)
        {
            // A line that falls outside its cell is not drawn: the cell does not grow, and a
            // reader must not find text the page does not show.
            if (top + appearance.LineHeight > cell.Bottom + 0.01 || top < cell.Y - 0.01)
            {
                top += appearance.LineHeight;
                continue;
            }

            double measured = appearance.Measure(line);
            double x = text.Alignment switch
            {
                HorizontalAlignment.Right => cell.Right - Padding - measured,
                HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous =>
                    cell.X + ((cell.Width - measured) / 2),
                _ => cell.X + Padding + (alignment.Indent * IndentStep),
            };

            _page.Items.Add(new TextItem
            {
                X = x,
                Y = top,
                Text = line,
                Appearance = appearance,
                Width = measured,
                Height = appearance.LineHeight,
                Clip = (cell.X, cell.Y, cell.Right, cell.Bottom),
                Tag = tag,
            });

            top += appearance.LineHeight;
        }
    }

    /// <summary>
    /// Breaks text into the lines a cell's width allows: at spaces where it can, and inside a
    /// word only when the word is wider than the cell and there is nowhere else to break.
    /// </summary>
    /// <param name="text">The text, in logical order.</param>
    /// <param name="appearance">How it is drawn, which is what decides its width.</param>
    /// <param name="room">How wide the lines may be.</param>
    /// <param name="alignment">The cell's alignment, whose reading order the lines follow.</param>
    private static List<string> Break(string text, TextAppearance appearance, double room, Alignment alignment)
    {
        List<string> lines = [];

        foreach (string paragraph in text.Split('\n'))
        {
            string rest = paragraph.TrimEnd('\r');

            while (rest.Length > 0)
            {
                int take = Fit(rest, appearance, room);
                lines.Add(Visual(rest[..take].TrimEnd(), alignment));
                rest = rest[take..].TrimStart();
            }

            if (paragraph.Length == 0)
                lines.Add("");
        }

        return lines.Count == 0 ? [""] : lines;
    }

    /// <summary>How many characters of the text fit one line.</summary>
    private static int Fit(string text, TextAppearance appearance, double room)
    {
        if (appearance.Measure(text) <= room)
            return text.Length;

        int lastSpace = -1;
        for (int i = 1; i <= text.Length; i++)
        {
            if (appearance.Measure(text[..i]) > room)
            {
                // Back up to the last place a reader would expect a break; with none, the word
                // is wider than the cell and is cut where it stopped fitting.
                return lastSpace > 0 ? lastSpace : Math.Max(1, i - 1);
            }

            if (i < text.Length && text[i] == ' ')
                lastSpace = i;
        }

        return text.Length;
    }

    /// <summary>
    /// Draws text turned on its side. The lines run along the cell's height and the position is
    /// the visual top-left of the strip they occupy, which is what the painter's one matrix
    /// carries them onto.
    /// </summary>
    private void Turned(CellText text, Alignment alignment, TextAppearance appearance, CellBox cell, TagRef? tag)
    {
        string shown = Visual(text.Text, alignment);

        if (alignment.TextRotation == Alignment.StackedRotation)
        {
            Stacked(shown, appearance, cell, tag);
            return;
        }

        // The format counts anticlockwise from the baseline, 1 to 90, and puts the clockwise
        // angles at 91 to 180 — 91 being one degree down. Anything else is drawn upright.
        int degrees = alignment.TextRotation;
        int rotation = degrees switch
        {
            >= 1 and <= 90 => 270,
            > 90 and <= 180 => 90,
            _ => 0,
        };

        if (rotation == 0)
        {
            Place(text.Alignment, alignment, appearance, shown, appearance.Measure(shown), cell, cell.X, cell.Right, tag);
            return;
        }

        if (degrees is not (90 or 180))
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.LayoutApproximated,
                "Text turned by an angle other than a quarter turn is drawn at the nearer quarter turn.",
                "rotation");
        }

        double measured = appearance.Measure(shown);
        double strip = appearance.LineHeight;

        _page.Items.Add(new TextItem
        {
            X = cell.X + ((cell.Width - strip) / 2),
            Y = rotation == 90 ? cell.Y + Padding : cell.Bottom - Padding - measured,
            Text = shown,
            Appearance = appearance,
            Width = measured,
            Height = strip,
            Rotation = rotation,
            Clip = (cell.X, cell.Y, cell.Right, cell.Bottom),
            Tag = tag,
        });
    }

    /// <summary>
    /// Draws text stacked: upright letters one under another, which is what a rotation of 255
    /// means. Each letter is its own line, so nothing has to be turned at all.
    /// </summary>
    private void Stacked(string text, TextAppearance appearance, CellBox cell, TagRef? tag)
    {
        double top = cell.Y + Padding;

        foreach (char c in text)
        {
            if (top + appearance.LineHeight > cell.Bottom)
                break;

            string letter = c.ToString();
            double measured = appearance.Measure(letter);

            _page.Items.Add(new TextItem
            {
                X = cell.X + ((cell.Width - measured) / 2),
                Y = top,
                Text = letter,
                Appearance = appearance,
                Width = measured,
                Height = appearance.LineHeight,
                Tag = tag,
            });

            top += appearance.LineHeight;
        }
    }
}
