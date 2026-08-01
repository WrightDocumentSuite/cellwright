using Cellwright.Styles;
using Inkwright.Text;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// The text inside the cells: what each shows, where in its cell it sits, and what happens when
/// it does not fit.
/// </summary>
internal sealed partial class CellPlacer
{
    /// <summary>What Excel leaves between a cell's edge and its text, in points.</summary>
    /// <remarks>
    /// Taken from the measurer rather than stated again: a cell is sized to hold its text before
    /// it is drawn, and a painter that padded differently would cut the text it was sized for.
    /// </remarks>
    private const double Padding = CellMeasurer.Padding;

    /// <summary>How far one indent step moves the text, in points — three digits of the font.</summary>
    private const double IndentStep = CellMeasurer.IndentStep;

    /// <summary>Draws the text of every cell on the page.</summary>
    public void Text()
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            for (int c = 0; c < _columns.Count; c++)
            {
                int row = _rows.Lines[r];
                int column = _columns.Lines[c];
                if (_covered.Contains((row, column)))
                    continue;

                CellValue value = _sheet.Values.Get(row, column);
                if (value.IsBlank)
                    continue;

                // A rule that shows only its icon or its bar means the number itself is not drawn.
                if (HidesValue(row, column))
                    continue;

                CellStyle style = StyleOf(row, column);
                TextAppearance appearance = _grid.Appearances.Of(style.Font);
                CellText text = CellContent.Of(_context, value, style, appearance.Color);
                if (text.IsEmpty)
                    continue;

                (double x, double y, double width, double height) = Box(c, r, row, column);

                // An icon sits at the left edge, and the text starts after it rather than under it.
                double icon = IconRoom(row, column, height);
                var cell = new CellBox(x + icon, y, width - icon, height, row, c);
                Draw(text, style, appearance with { Color = text.Color }, cell, Tag(row, column));
            }
        }
    }

    /// <summary>Where one cell sits on the page, and which line of the sheet it came from.</summary>
    private readonly record struct CellBox(
        double X, double Y, double Width, double Height, int Row, int Slot)
    {
        public double Right => X + Width;

        public double Bottom => Y + Height;
    }

    /// <summary>
    /// Draws one cell's text: turned, wrapped, shrunk or spilled, whichever its style and its
    /// width call for.
    /// </summary>
    private void Draw(CellText text, CellStyle style, TextAppearance appearance, CellBox cell, TagRef? tag)
    {
        Alignment alignment = style.Alignment;
        double room = CellMeasurer.Room(cell.Width, alignment);

        if (alignment.TextRotation != 0)
        {
            Turned(text, alignment, appearance, cell, tag);
            return;
        }

        string shown = Visual(text.Text, alignment);
        double measured = appearance.Measure(shown);

        if (alignment.WrapText && measured > room)
        {
            Wrapped(text, alignment, appearance, cell, room, tag);
            return;
        }

        if (alignment.ShrinkToFit && measured > room)
        {
            appearance = Shrunk(appearance, room, measured);
            measured = appearance.Measure(shown);
        }

        // A number that does not fit its column is not shown wrong: Excel fills the cell with
        // hashes instead, which is a reader's cue that the column has to be widened.
        if (text.IsNumeric && measured > room + 0.01)
        {
            Hashes(appearance, cell, room, tag);
            return;
        }

        (double left, double right) = Reach(cell, alignment, text.Alignment, measured);
        Place(text.Alignment, alignment, appearance, shown, measured, cell, left, right, tag);
    }

    /// <summary>Puts one line of text where its alignments say, clipped to the room it may use.</summary>
    private void Place(
        HorizontalAlignment horizontal,
        Alignment alignment,
        TextAppearance appearance,
        string shown,
        double measured,
        CellBox cell,
        double reachLeft,
        double reachRight,
        TagRef? tag)
    {
        double x = horizontal switch
        {
            HorizontalAlignment.Right => cell.Right - Padding - measured,
            HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous =>
                cell.X + ((cell.Width - measured) / 2),
            _ => cell.X + Padding + (alignment.Indent * IndentStep),
        };

        bool clipped = x < reachLeft - 0.01 || x + measured > reachRight + 0.01;
        if (clipped)
        {
            // What the reader will not see is not drawn at all. Clipping alone would hide it,
            // but a glyph hidden behind a clip is still in the file: it would be found by a
            // search and read out by a reader, neither of which is what the page shows.
            (shown, x) = Trim(shown, appearance, x, reachLeft, reachRight);
            measured = appearance.Measure(shown);

            if (shown.Length == 0)
                return;
        }

        _page.Items.Add(new TextItem
        {
            X = x,
            Y = Down(alignment.Vertical, cell.Y, cell.Height, appearance.LineHeight),
            Text = shown,
            Appearance = appearance,
            Width = measured,
            Height = appearance.LineHeight,
            Clip = clipped ? (reachLeft, cell.Y, reachRight, cell.Bottom) : null,
            Tag = tag,
        });
    }

    /// <summary>
    /// Cuts a line down to the characters that fall inside the room it may use, and says where
    /// what is left begins. A character straddling the edge is kept and the clip cuts it in half,
    /// which is what Excel shows at the edge of a column.
    /// </summary>
    private static (string Text, double X) Trim(
        string text, TextAppearance appearance, double x, double left, double right)
    {
        int from = 0;
        while (from < text.Length && x + appearance.Measure(text[..(from + 1)]) <= left)
            from++;

        double start = x + (from > 0 ? appearance.Measure(text[..from]) : 0);
        int to = from;
        while (to < text.Length && start + appearance.Measure(text[from..(to + 1)]) <= right + 0.01)
            to++;

        // One character over the edge is kept so that the clip has something to cut.
        if (to < text.Length)
            to++;

        return (text[from..Math.Max(from, to)], start);
    }

    /// <summary>
    /// How far the text may reach. Text without wrapping spills into the empty cells beside it —
    /// to the right when it is left-aligned, to the left when it is right-aligned, both ways
    /// when it is centred — and stops at the first neighbour that holds anything, where Excel
    /// clips it.
    /// </summary>
    private (double Left, double Right) Reach(
        CellBox cell, Alignment alignment, HorizontalAlignment horizontal, double measured)
    {
        if (alignment.WrapText || alignment.ShrinkToFit || measured <= cell.Width - (2 * Padding) + 0.01)
            return (cell.X, cell.Right);

        double left = cell.X;
        double right = cell.Right;

        if (horizontal != HorizontalAlignment.Right)
            right += Free(cell, step: 1);

        if (horizontal != HorizontalAlignment.Left)
            left -= Free(cell, step: -1);

        return (left, right);
    }

    /// <summary>
    /// How much empty room one direction offers, in points. Past the last printed column there
    /// is nothing in the way at all, so the text runs to the edge of the page.
    /// </summary>
    private double Free(CellBox cell, int step)
    {
        double room = 0;
        int i = cell.Slot + step;

        for (; i >= 0 && i < _columns.Count; i += step)
        {
            int column = _columns.Lines[i];
            if (!_sheet.Values.Get(cell.Row, column).IsBlank || _covered.Contains((cell.Row, column)))
                return room;

            room += _columns.Size[i];
        }

        // The walk fell off the end of the grid: what is left is the page's own margin.
        return step > 0
            ? room + Math.Max(0, PageReach - _columns.Far(_columns.Count - 1))
            : room + Math.Max(0, _columns.Near[0]);
    }

    /// <summary>Where the top of a line box sits down its cell.</summary>
    private static double Down(VerticalAlignment alignment, double y, double height, double line) => alignment switch
    {
        VerticalAlignment.Top => y + Padding,
        VerticalAlignment.Center => y + ((height - line) / 2),
        VerticalAlignment.Justify or VerticalAlignment.Distributed => y + Padding,
        _ => y + height - Padding - line,
    };

    /// <summary>Fills a cell with hashes, as Excel does with a number too wide for its column.</summary>
    private void Hashes(TextAppearance appearance, CellBox cell, double room, TagRef? tag)
    {
        double one = Math.Max(0.1, appearance.Measure("#"));
        int count = Math.Max(1, (int)(room / one));
        string hashes = new('#', count);

        _page.Items.Add(new TextItem
        {
            X = cell.X + Padding,
            Y = Down(VerticalAlignment.Bottom, cell.Y, cell.Height, appearance.LineHeight),
            Text = hashes,
            Appearance = appearance,
            Width = one * count,
            Height = appearance.LineHeight,
            Tag = tag,
        });
    }

    /// <summary>The same appearance at the size that makes the text fit.</summary>
    private static TextAppearance Shrunk(TextAppearance appearance, double room, double measured)
    {
        double scale = Math.Clamp(room / measured, 0.1, 1);
        return appearance with
        {
            Size = appearance.Size * scale,
            Ascent = appearance.Ascent * scale,
            Descent = appearance.Descent * scale,
        };
    }

    /// <summary>
    /// The text in the order it is painted. A cell whose reading order is right-to-left, or
    /// whose text is Hebrew or Arabic, is reordered and shaped; everything else passes through.
    /// </summary>
    private static string Visual(string text, Alignment alignment)
    {
        bool rightToLeft = alignment.ReadingOrder == ReadingOrder.RightToLeft;
        return rightToLeft || BidiLayout.HasRightToLeft(text)
            ? BidiLayout.ToVisualLine(text, rightToLeft)
            : text;
    }
}
