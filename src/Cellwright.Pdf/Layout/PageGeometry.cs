namespace Cellwright.Pdf.Layout;

/// <summary>
/// The measurements of one page in points, with the origin at the top-left corner and the
/// vertical axis pointing down.
/// </summary>
/// <remarks>
/// A spreadsheet measures its margins in inches from the paper edge; PDF measures in points from
/// the bottom. Rather than flipping in a dozen places, the whole layout runs in this top-down
/// space and the renderer flips once, at the moment it writes a coordinate.
/// <see cref="ToPdfY"/> is that one place.
/// </remarks>
internal readonly record struct PageGeometry
{
    private PageGeometry(
        double width,
        double height,
        double left,
        double right,
        double top,
        double bottom,
        double header,
        double footer)
    {
        Width = width;
        Height = height;
        MarginLeft = left;
        MarginRight = right;
        MarginTop = top;
        MarginBottom = bottom;
        HeaderDistance = header;
        FooterDistance = footer;
    }

    /// <summary>The page width in points.</summary>
    public double Width { get; }

    /// <summary>The page height in points.</summary>
    public double Height { get; }

    /// <summary>Space to the left of the printed grid.</summary>
    public double MarginLeft { get; }

    /// <summary>Space to the right of it.</summary>
    public double MarginRight { get; }

    /// <summary>Space above it, from the paper edge.</summary>
    public double MarginTop { get; }

    /// <summary>Space below it, to the paper edge.</summary>
    public double MarginBottom { get; }

    /// <summary>How far the header sits from the top of the paper.</summary>
    public double HeaderDistance { get; }

    /// <summary>How far the footer sits from the bottom of the paper.</summary>
    public double FooterDistance { get; }

    /// <summary>The left edge of the printed area.</summary>
    public double ContentLeft => MarginLeft;

    /// <summary>Its right edge.</summary>
    public double ContentRight => Width - MarginRight;

    /// <summary>Its top edge.</summary>
    public double ContentTop => MarginTop;

    /// <summary>Its bottom edge.</summary>
    public double ContentBottom => Height - MarginBottom;

    /// <summary>How wide it is.</summary>
    public double ContentWidth => Math.Max(1, ContentRight - ContentLeft);

    /// <summary>How tall it is.</summary>
    public double ContentHeight => Math.Max(1, ContentBottom - ContentTop);

    /// <summary>Turns a top-down coordinate into the bottom-up one a PDF is written in.</summary>
    /// <param name="y">Points down from the top of the page.</param>
    public double ToPdfY(double y) => Height - y;

    /// <summary>Reads the geometry a sheet's print settings ask for.</summary>
    /// <param name="setup">The sheet's print settings.</param>
    public static PageGeometry From(PageSetup setup)
    {
        ArgumentNullException.ThrowIfNull(setup);

        (double paperWidth, double paperHeight) = SheetUnits.Paper(setup.PaperSize);
        bool landscape = setup.Orientation == PageOrientation.Landscape;
        double width = landscape ? paperHeight : paperWidth;
        double height = landscape ? paperWidth : paperHeight;

        PageMargins margins = setup.Margins ?? PageMargins.Normal;

        // A margin wider than the paper would leave nothing to print on; the pair is squeezed
        // back to leave an inch, which is what a printer driver does with the same nonsense.
        (double left, double right) = Fit(margins.Left, margins.Right, width);
        (double top, double bottom) = Fit(margins.Top, margins.Bottom, height);

        return new PageGeometry(
            width,
            height,
            SheetUnits.Points(left),
            SheetUnits.Points(right),
            SheetUnits.Points(top),
            SheetUnits.Points(bottom),
            SheetUnits.Points(Math.Max(0, margins.Header)),
            SheetUnits.Points(Math.Max(0, margins.Footer)));
    }

    private static (double Near, double Far) Fit(double near, double far, double extent)
    {
        near = Math.Max(0, near);
        far = Math.Max(0, far);

        double room = extent / SheetUnits.PointsPerInch;
        if (near + far <= room - 1)
            return (near, far);

        double scale = Math.Max(0, room - 1) / Math.Max(0.001, near + far);
        return (near * scale, far * scale);
    }
}
