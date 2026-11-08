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

                (double x, double y, double width, double height) = Box(c, r, row, column);
                CellValue value = _sheet.Values.Get(row, column);
                if (!value.IsBlank && !HidesValue(row, column))
                {
                    CellStyle style = StyleOf(row, column);
                    TextAppearance appearance = _grid.Appearances.Of(style.Font);
                    CellText text = CellContent.Of(_context, value, style, appearance.Color);
                    if (!text.IsEmpty)
                    {
                        // An icon sits at the left edge, and the text starts after it rather than under it.
                        double icon = IconRoom(row, column, height);
                        var box = new CellBox(x + icon, y, width - icon, height, row, c);
                        RichText? rich = _sheet.Cell(row, column).RichText;
                        if (rich is { IsFormatted: true } &&
                            style.Alignment.TextRotation == 0 &&
                            string.Equals(text.Text, rich.Text, StringComparison.Ordinal))
                        {
                            DrawRich(rich, text.Alignment, style, box, Tag(row, column));
                        }
                        else
                        {
                            Draw(text, style, appearance with { Color = text.Color }, box, Tag(row, column));
                        }
                    }
                }

                Comments?.AddAt(
                    _page,
                    new CellRef(row, column),
                    x,
                    y,
                    width,
                    height,
                    Tag(row, column));
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

    private readonly record struct RichSegment(string Text, TextAppearance Appearance);

    private readonly record struct RichLine(
        List<RichSegment> Segments, double Width, double Ascent, double LineHeight);

    /// <summary>
    /// Draws a cell whose text carries per-run character formatting (<see cref="Cell.RichText"/>).
    /// Each run is drawn in its own <see cref="TextAppearance"/> while sharing a common baseline
    /// per line and obeying the cell's horizontal/vertical alignment, wrapping, shrink-to-fit,
    /// and overflow clipping.
    /// </summary>
    private void DrawRich(
        RichText rich, HorizontalAlignment horizontal, CellStyle style, CellBox cell, TagRef? tag)
    {
        Alignment alignment = style.Alignment;
        double room = CellMeasurer.Room(cell.Width, alignment);
        List<RichLine> lines = BuildRichLines(rich, style.Font, _grid.Appearances, alignment, room);
        if (lines.Count == 0)
            return;

        if (!alignment.WrapText && lines.Count == 1)
        {
            RichLine line = lines[0];
            if (alignment.ShrinkToFit && line.Width > room)
                line = ShrinkRichLine(line, room);

            (double reachLeft, double reachRight) = Reach(cell, alignment, horizontal, line.Width);
            double x = horizontal switch
            {
                HorizontalAlignment.Right => cell.Right - Padding - line.Width,
                HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous =>
                    cell.X + ((cell.Width - line.Width) / 2),
                _ => cell.X + Padding + (alignment.Indent * IndentStep),
            };
            double top = Down(alignment.Vertical, cell.Y, cell.Height, line.LineHeight);
            PlaceRichLine(line, x, top, cell, reachLeft, reachRight, tag);
            return;
        }

        double totalHeight = 0;
        foreach (RichLine line in lines)
            totalHeight += line.LineHeight;

        double cursorTop = Down(alignment.Vertical, cell.Y, cell.Height, totalHeight);
        foreach (RichLine line in lines)
        {
            if (cursorTop + line.LineHeight > cell.Bottom + 0.01 || cursorTop < cell.Y - 0.01)
            {
                cursorTop += line.LineHeight;
                continue;
            }

            double x = horizontal switch
            {
                HorizontalAlignment.Right => cell.Right - Padding - line.Width,
                HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous =>
                    cell.X + ((cell.Width - line.Width) / 2),
                _ => cell.X + Padding + (alignment.Indent * IndentStep),
            };
            PlaceRichLine(line, x, cursorTop, cell, cell.X, cell.Right, tag);
            cursorTop += line.LineHeight;
        }
    }

    private void PlaceRichLine(
        RichLine line,
        double startX,
        double lineTop,
        CellBox cell,
        double reachLeft,
        double reachRight,
        TagRef? tag)
    {
        double cursorX = startX;
        foreach (RichSegment segment in line.Segments)
        {
            if (segment.Text.Length == 0)
                continue;

            double segWidth = segment.Appearance.Measure(segment.Text);
            double segX = cursorX;
            cursorX += segWidth;

            string shown = segment.Text;
            bool clipped = segX < reachLeft - 0.01 || segX + segWidth > reachRight + 0.01;
            if (clipped)
            {
                (shown, segX) = Trim(shown, segment.Appearance, segX, reachLeft, reachRight);
                segWidth = segment.Appearance.Measure(shown);
                if (shown.Length == 0)
                    continue;
            }

            _page.Items.Add(new TextItem
            {
                X = segX,
                Y = lineTop + (line.Ascent - segment.Appearance.Ascent),
                Text = shown,
                Appearance = segment.Appearance,
                Width = segWidth,
                Height = segment.Appearance.LineHeight,
                Clip = clipped ? (reachLeft, cell.Y, reachRight, cell.Bottom) : null,
                Tag = tag,
            });
        }
    }

    private static RichLine ShrinkRichLine(RichLine line, double room)
    {
        var shrunk = new List<RichSegment>(line.Segments.Count);
        double width = 0;
        double ascent = 0;
        double lineHeight = 0;
        foreach (RichSegment seg in line.Segments)
        {
            TextAppearance scaled = Shrunk(seg.Appearance, room, line.Width);
            shrunk.Add(new RichSegment(seg.Text, scaled));
            width += scaled.Measure(seg.Text);
            ascent = Math.Max(ascent, scaled.Ascent);
            lineHeight = Math.Max(lineHeight, scaled.LineHeight);
        }

        return new RichLine(shrunk, width, ascent, lineHeight);
    }

    internal static List<(double Width, double LineHeight)> MeasureRichLines(
        RichText rich, FontStyle defaultFont, CellAppearances appearances, Alignment alignment, double room)
    {
        List<RichLine> lines = BuildRichLines(rich, defaultFont, appearances, alignment, room);
        var metrics = new List<(double Width, double LineHeight)>(lines.Count);
        foreach (RichLine line in lines)
            metrics.Add((line.Width, line.LineHeight));
        return metrics;
    }

    private static List<RichLine> BuildRichLines(
        RichText rich, FontStyle defaultFont, CellAppearances appearances, Alignment alignment, double room)
    {
        var paragraphs = new List<List<RichSegment>> { new() };
        TextAppearance baseAppearance = appearances.Of(defaultFont);

        foreach (RichTextRun run in rich)
        {
            if (run.Text.Length == 0)
                continue;

            TextAppearance runAppearance = run.Font is { } font ? appearances.Of(font) : baseAppearance;
            string[] parts = run.Text.Split('\n');
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0)
                    paragraphs.Add([]);

                string piece = parts[i].TrimEnd('\r');
                if (piece.Length > 0)
                    paragraphs[^1].Add(new RichSegment(Visual(piece, alignment), runAppearance));
            }
        }

        var result = new List<RichLine>();
        foreach (List<RichSegment> paragraph in paragraphs)
        {
            if (!alignment.WrapText || paragraph.Count == 0)
            {
                result.Add(ToRichLine(paragraph, baseAppearance));
                continue;
            }

            WrapRichParagraph(paragraph, baseAppearance, room, result);
        }

        return result;
    }

    private static void WrapRichParagraph(
        List<RichSegment> paragraph, TextAppearance fallback, double room, List<RichLine> output)
    {
        var chars = new List<(char Char, TextAppearance Appearance)>();
        foreach (RichSegment seg in paragraph)
        {
            foreach (char ch in seg.Text)
                chars.Add((ch, seg.Appearance));
        }

        int pos = 0;
        while (pos < chars.Count)
        {
            int take = FitRich(chars, pos, room);
            int end = pos + take;
            while (end > pos && chars[end - 1].Char == ' ')
                end--;

            output.Add(SliceToRichLine(chars, pos, end, fallback));
            pos += take;
            while (pos < chars.Count && chars[pos].Char == ' ')
                pos++;
        }
    }

    private static int FitRich(List<(char Char, TextAppearance Appearance)> chars, int start, double room)
    {
        int remaining = chars.Count - start;
        if (MeasureSlice(chars, start, remaining) <= room)
            return remaining;

        int lastSpace = -1;
        for (int len = 1; len <= remaining; len++)
        {
            if (MeasureSlice(chars, start, len) > room)
                return lastSpace > 0 ? lastSpace : Math.Max(1, len - 1);

            if (chars[start + len - 1].Char == ' ')
                lastSpace = len - 1;
        }

        return remaining;
    }

    private static double MeasureSlice(List<(char Char, TextAppearance Appearance)> chars, int start, int length)
    {
        if (length <= 0)
            return 0;

        double total = 0;
        int i = start;
        int end = start + length;
        while (i < end)
        {
            TextAppearance current = chars[i].Appearance;
            int runStart = i;
            i++;
            while (i < end && ReferenceEquals(chars[i].Appearance, current))
                i++;

            var buffer = new char[i - runStart];
            for (int k = 0; k < buffer.Length; k++)
                buffer[k] = chars[runStart + k].Char;

            total += current.Measure(new string(buffer));
        }

        return total;
    }

    private static RichLine SliceToRichLine(
        List<(char Char, TextAppearance Appearance)> chars, int start, int end, TextAppearance fallback)
    {
        var segments = new List<RichSegment>();
        int i = start;
        while (i < end)
        {
            TextAppearance current = chars[i].Appearance;
            int runStart = i;
            i++;
            while (i < end && ReferenceEquals(chars[i].Appearance, current))
                i++;

            var buffer = new char[i - runStart];
            for (int k = 0; k < buffer.Length; k++)
                buffer[k] = chars[runStart + k].Char;

            segments.Add(new RichSegment(new string(buffer), current));
        }

        return ToRichLine(segments, fallback);
    }

    private static RichLine ToRichLine(List<RichSegment> segments, TextAppearance fallback)
    {
        if (segments.Count == 0)
            return new RichLine(segments, 0, fallback.Ascent, fallback.LineHeight);

        double width = 0;
        double ascent = 0;
        double lineHeight = 0;
        foreach (RichSegment seg in segments)
        {
            width += seg.Appearance.Measure(seg.Text);
            ascent = Math.Max(ascent, seg.Appearance.Ascent);
            lineHeight = Math.Max(lineHeight, seg.Appearance.LineHeight);
        }

        return new RichLine(segments, width, ascent, lineHeight);
    }
}
