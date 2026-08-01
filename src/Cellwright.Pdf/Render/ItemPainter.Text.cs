using Cellwright.Pdf.Layout;
using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Render;

/// <summary>Drawing text: the glyphs themselves and the rules through and under them.</summary>
internal sealed partial class ItemPainter
{
    private void PaintText(TextItem item)
    {
        if (item.Text.Length == 0)
            return;

        bool clipped = item.Clip is not null;
        if (clipped)
            Clip(item.Clip!.Value);

        if (item.Rotation != 0)
            PaintTurned(item);
        else
            PaintUpright(item, item.X, item.Y + item.Appearance.Ascent);

        if (clipped)
            _canvas.Restore();
    }

    /// <summary>Cuts everything that follows down to a rectangle, until the state is restored.</summary>
    private void Clip((double Left, double Top, double Right, double Bottom) box)
    {
        _canvas.Save()
               .Rectangle(
                   box.Left,
                   _geometry.ToPdfY(box.Bottom),
                   Math.Max(0, box.Right - box.Left),
                   Math.Max(0, box.Bottom - box.Top))
               .Clip();
    }

    private void PaintUpright(TextItem item, double x, double baseline)
    {
        TextAppearance appearance = item.Appearance;

        int mcid = BeginTag(item.Tag);
        _canvas.Save()
               .BeginText()
               .FillColor(appearance.Color)
               .Font(appearance.Font, appearance.Size)
               .TextMatrix(PdfMatrix.Translation(x, _geometry.ToPdfY(baseline) + appearance.Rise))
               .ShowText(item.Text)
               .EndText()
               .Restore();
        EndTag(item.Tag, mcid);

        Decorate(item, x, baseline);
    }

    /// <summary>
    /// Draws text turned on its side. The turned frame imitates the page — same height, top-down
    /// coordinates — so the upright painter works inside it unchanged, and one matrix carries the
    /// whole frame onto the strip the text occupies.
    /// </summary>
    private void PaintTurned(TextItem item)
    {
        double page = _geometry.Height;
        _canvas.Save();

        if (item.Rotation == 90)
        {
            // Read downwards, glyph tops to the right: the frame turns clockwise about the
            // strip's top-left corner.
            _canvas.Transform(PdfMatrix.Translation(item.X + item.Height - page, page - item.Y));
            _canvas.Transform(PdfMatrix.Rotation(-90));
        }
        else
        {
            // Read upwards, glyph tops to the left: anticlockwise, from the strip's bottom.
            _canvas.Transform(PdfMatrix.Translation(item.X + page, page - item.Y - item.Width));
            _canvas.Transform(PdfMatrix.Rotation(90));
        }

        PaintUpright(item, 0, item.Appearance.Ascent);
        _canvas.Restore();
    }

    /// <summary>Rules the text under and through, as its font style asks.</summary>
    private void Decorate(TextItem item, double x, double baseline)
    {
        TextAppearance appearance = item.Appearance;
        if (item.Width <= 0)
            return;

        if (appearance.Underline != UnderlineStyle.None)
        {
            double thickness = Math.Max(0.4, appearance.Size * UnderlineWeight);
            double y = baseline + (appearance.Size * UnderlineOffset) - appearance.Rise;
            Rule(x, y, x + item.Width, thickness, appearance.Color);

            if (appearance.Underline is UnderlineStyle.Double or UnderlineStyle.DoubleAccounting)
                Rule(x, y + (thickness * 2), x + item.Width, thickness, appearance.Color);
        }

        if (appearance.Strikethrough)
        {
            double y = baseline - (appearance.Size * StrikeOffset) - appearance.Rise;
            Rule(x, y, x + item.Width, Math.Max(0.4, appearance.Size * UnderlineWeight), appearance.Color);
        }
    }

    private void Rule(double x1, double y, double x2, double thickness, PdfColor color) =>
        PaintStroke(new StrokeItem
        {
            X = x1,
            Y = y,
            X2 = x2,
            Y2 = y,
            Thickness = thickness,
            Color = color,
        });
}
