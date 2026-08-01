using Cellwright.Pdf.Layout;
using Inkwright;

namespace Cellwright.Pdf.Render;

/// <summary>
/// The icons an icon-set rule puts in its cells.
/// </summary>
/// <remarks>
/// Drawn as outlines rather than as glyphs of a symbol font, because the font that carries the
/// right glyphs is not one every machine has, and a missing glyph would be a blank where the
/// sheet says something. A shape is always drawable.
/// </remarks>
internal sealed partial class ItemPainter
{
    /// <summary>How much of the icon's square the shape actually occupies.</summary>
    private const double IconInset = 0.12;

    private void PaintIcon(IconItem icon)
    {
        if (icon.Size <= 0)
            return;

        double inset = icon.Size * IconInset;
        double side = icon.Size - (2 * inset);
        if (side <= 0)
            return;

        double left = icon.X + inset;

        // The layout works top-down and the page is bottom-up, so the box is flipped once here
        // and every shape below is drawn in page space with y rising.
        double bottom = _geometry.ToPdfY(icon.Y + icon.Size - inset);

        _canvas.Save().FillColor(icon.Color).StrokeColor(icon.Color);

        switch (icon.Shape)
        {
            case IconShape.Circle:
                _canvas.Ellipse(PdfRectangle.FromSize(left, bottom, side, side)).Fill();
                break;

            case IconShape.Arrow:
                Arrow(left, bottom, side, icon.Angle);
                break;

            case IconShape.Flag:
                Flag(left, bottom, side);
                break;

            case IconShape.Triangle:
                Polygon([
                    (left + (side / 2), bottom + side),
                    (left + side, bottom),
                    (left, bottom)]);
                break;

            case IconShape.Diamond:
                Polygon([
                    (left + (side / 2), bottom + side),
                    (left + side, bottom + (side / 2)),
                    (left + (side / 2), bottom),
                    (left, bottom + (side / 2))]);
                break;

            case IconShape.Check:
                Check(left, bottom, side);
                break;

            case IconShape.Cross:
                Cross(left, bottom, side);
                break;

            case IconShape.Exclamation:
                Exclamation(left, bottom, side);
                break;

            case IconShape.Rating:
                Rating(left, bottom, side, icon.Fraction);
                break;

            case IconShape.Quarters:
                Quarters(left, bottom, side, icon.Fraction);
                break;

            default:
                break;
        }

        _canvas.Restore();
    }

    /// <summary>An arrow of the given side, turned clockwise from pointing up.</summary>
    private void Arrow(double left, double bottom, double side, double degrees)
    {
        double centreX = left + (side / 2);
        double centreY = bottom + (side / 2);
        double radians = degrees * Math.PI / 180;

        // The outline is written once, pointing up, around the origin; turning it is arithmetic
        // rather than a matrix so that the path stays one list of points.
        (double X, double Y)[] outline =
        [
            (0, 0.5),
            (0.35, 0.1),
            (0.15, 0.1),
            (0.15, -0.5),
            (-0.15, -0.5),
            (-0.15, 0.1),
            (-0.35, 0.1),
        ];

        var points = new (double X, double Y)[outline.Length];
        for (int i = 0; i < outline.Length; i++)
        {
            double x = outline[i].X * side;
            double y = outline[i].Y * side;
            points[i] = (
                centreX + (x * Math.Cos(radians)) + (y * Math.Sin(radians)),
                centreY - (x * Math.Sin(radians)) + (y * Math.Cos(radians)));
        }

        Polygon(points);
    }

    private void Flag(double left, double bottom, double side)
    {
        // A staff up the left, and a pennant off the top of it.
        _canvas.Rectangle(left + (side * 0.1), bottom, side * 0.12, side).Fill();
        Polygon([
            (left + (side * 0.22), bottom + side),
            (left + side, bottom + (side * 0.78)),
            (left + (side * 0.22), bottom + (side * 0.55))]);
    }

    private void Check(double left, double bottom, double side)
    {
        _canvas.LineWidth(side * 0.18)
               .MoveTo(left + (side * 0.12), bottom + (side * 0.52))
               .LineTo(left + (side * 0.4), bottom + (side * 0.22))
               .LineTo(left + (side * 0.9), bottom + (side * 0.8))
               .Stroke();
    }

    private void Cross(double left, double bottom, double side)
    {
        _canvas.LineWidth(side * 0.18)
               .MoveTo(left + (side * 0.15), bottom + (side * 0.15))
               .LineTo(left + (side * 0.85), bottom + (side * 0.85))
               .Stroke()
               .MoveTo(left + (side * 0.15), bottom + (side * 0.85))
               .LineTo(left + (side * 0.85), bottom + (side * 0.15))
               .Stroke();
    }

    private void Exclamation(double left, double bottom, double side)
    {
        double width = side * 0.2;
        double centre = left + ((side - width) / 2);

        _canvas.Rectangle(centre, bottom + (side * 0.3), width, side * 0.7).Fill();
        _canvas.Rectangle(centre, bottom, width, side * 0.18).Fill();
    }

    /// <summary>Bars of rising height, the lower ones filled to show the rating.</summary>
    private void Rating(double left, double bottom, double side, double fraction)
    {
        const int Bars = 4;
        double gap = side * 0.08;
        double width = (side - (gap * (Bars - 1))) / Bars;
        int filled = (int)Math.Round(Math.Clamp(fraction, 0, 1) * Bars);

        for (int i = 0; i < Bars; i++)
        {
            double height = side * (0.35 + (0.65 * (i + 1) / Bars));
            double x = left + (i * (width + gap));

            _canvas.Rectangle(x, bottom, width, height);
            if (i < filled)
                _canvas.Fill();
            else
                _canvas.LineWidth(side * 0.05).Stroke();
        }
    }

    /// <summary>A disc with a share of it filled, drawn clockwise from the top.</summary>
    private void Quarters(double left, double bottom, double side, double fraction)
    {
        double radius = side / 2;
        double centreX = left + radius;
        double centreY = bottom + radius;

        _canvas.LineWidth(side * 0.08).Ellipse(PdfRectangle.FromSize(left, bottom, side, side)).Stroke();

        double share = Math.Clamp(fraction, 0, 1);
        if (share <= 0)
            return;

        if (share >= 1)
        {
            _canvas.Ellipse(PdfRectangle.FromSize(left, bottom, side, side)).Fill();
            return;
        }

        // Approximated by a fan of short segments: a wedge is a handful of them, and the icon is
        // a few points across, where the difference from a true arc is well under a pixel.
        const int Steps = 24;
        int used = Math.Max(1, (int)Math.Round(Steps * share));

        _canvas.MoveTo(centreX, centreY);
        for (int i = 0; i <= used; i++)
        {
            double angle = (Math.PI / 2) - (2 * Math.PI * share * i / used);
            _canvas.LineTo(centreX + (radius * Math.Cos(angle)), centreY + (radius * Math.Sin(angle)));
        }

        _canvas.ClosePath().Fill();
    }

    private void Polygon(scoped ReadOnlySpan<(double X, double Y)> points)
    {
        if (points.Length == 0)
            return;

        _canvas.MoveTo(points[0].X, points[0].Y);
        for (int i = 1; i < points.Length; i++)
            _canvas.LineTo(points[i].X, points[i].Y);

        _canvas.ClosePath().Fill();
    }
}
