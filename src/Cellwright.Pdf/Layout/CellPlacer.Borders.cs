using Cellwright.Styles;
using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// The lines a cell's own style draws around it, which sit over whatever the grid drew.
/// </summary>
/// <remarks>
/// Two cells share an edge and both may name a border for it. Excel does not merge them: it
/// draws the one belonging to the cell it is drawing, so the later cell wins. Drawing both in
/// order comes to the same picture and needs no rule about which is heavier.
/// </remarks>
internal sealed partial class CellPlacer
{
    /// <summary>Draws the borders of every cell on the page.</summary>
    public void Borders()
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            for (int c = 0; c < _columns.Count; c++)
            {
                int row = _rows.Lines[r];
                int column = _columns.Lines[c];
                if (_covered.Contains((row, column)))
                    continue;

                BorderStyle border = StyleOf(row, column).Border;
                if (border == BorderStyle.None)
                    continue;

                (double x, double y, double width, double height) = Box(c, r, row, column);
                Edges(border, x, y, width, height);
            }
        }
    }

    private void Edges(BorderStyle border, double x, double y, double width, double height)
    {
        Edge(border.Top, x, y, x + width, y);
        Edge(border.Bottom, x, y + height, x + width, y + height);
        Edge(border.Left, x, y, x, y + height);
        Edge(border.Right, x + width, y, x + width, y + height);

        if (border.DiagonalDown)
            Edge(border.Diagonal, x, y, x + width, y + height);

        if (border.DiagonalUp)
            Edge(border.Diagonal, x, y + height, x + width, y);
    }

    private void Edge(BorderEdge edge, double x1, double y1, double x2, double y2)
    {
        if (edge.Line == BorderLineStyle.None)
            return;

        _page.Items.Add(new StrokeItem
        {
            X = x1,
            Y = y1,
            X2 = x2,
            Y2 = y2,
            Thickness = Thickness(edge.Line),
            Color = _context.ColorOf(edge.Color, PdfColor.Black),
            Style = edge.Line,
        });
    }

    /// <summary>
    /// How thick a border line is drawn. The format names weights rather than measurements, and
    /// these are the widths Excel prints them at.
    /// </summary>
    private static double Thickness(BorderLineStyle line) => line switch
    {
        BorderLineStyle.Hair => 0.25,
        BorderLineStyle.Thin or BorderLineStyle.Dotted or BorderLineStyle.Dashed
            or BorderLineStyle.DashDot or BorderLineStyle.DashDotDot => 0.5,
        BorderLineStyle.Medium or BorderLineStyle.MediumDashed or BorderLineStyle.MediumDashDot
            or BorderLineStyle.MediumDashDotDot or BorderLineStyle.SlantDashDot => 1.5,
        BorderLineStyle.Thick => 2.5,
        BorderLineStyle.Double => 1,
        _ => 0.5,
    };
}
