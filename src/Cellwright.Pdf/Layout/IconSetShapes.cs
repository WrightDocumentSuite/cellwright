using Inkwright;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// Which outline and colour each icon of each built-in set is drawn with.
/// </summary>
/// <remarks>
/// The seventeen sets the format names (ST_IconSetType, §18.18.42) are seven shapes between
/// them, and the colour is nearly always a ramp from the worst band to the best. Naming the ramp
/// and the shape separately is what keeps this a table rather than seventeen drawing routines.
/// </remarks>
internal static class IconSetShapes
{
    private static readonly PdfColor Red = PdfColor.FromRgb(0xFF0000);
    private static readonly PdfColor Amber = PdfColor.FromRgb(0xFFC000);
    private static readonly PdfColor Green = PdfColor.FromRgb(0x00B050);
    private static readonly PdfColor Grey = PdfColor.FromRgb(0xA5A5A5);
    private static readonly PdfColor Black = PdfColor.FromRgb(0x404040);
    private static readonly PdfColor Pink = PdfColor.FromRgb(0xE8908A);
    private static readonly PdfColor Blue = PdfColor.FromRgb(0x5B9BD5);

    /// <summary>
    /// The icon a set shows for one band, or null when the set is not one this knows.
    /// </summary>
    /// <param name="set">Name of the set, as the file spells it.</param>
    /// <param name="index">Which band, counted from the lowest.</param>
    /// <param name="count">How many bands the set has.</param>
    public static (IconShape Shape, PdfColor Color, double Angle, double Fraction)? Of(
        string set, int index, int count)
    {
        int band = Math.Clamp(index, 0, Math.Max(count - 1, 0));
        double share = count <= 1 ? 1 : band / (double)(count - 1);

        return set switch
        {
            "3Arrows" or "4Arrows" or "5Arrows" =>
                (IconShape.Arrow, Ramp(share), ArrowAngle(share), 1),

            "3ArrowsGray" or "4ArrowsGray" or "5ArrowsGray" =>
                (IconShape.Arrow, Grey, ArrowAngle(share), 1),

            "3Flags" => (IconShape.Flag, Ramp(share), 0, 1),

            "3TrafficLights1" or "3TrafficLights2" => (IconShape.Circle, Ramp(share), 0, 1),

            // The four-light set puts black below red rather than stretching the ramp.
            "4TrafficLights" => (IconShape.Circle, band switch
            {
                0 => Black,
                1 => Red,
                2 => Amber,
                _ => Green,
            }, 0, 1),

            "4RedToBlack" => (IconShape.Circle, band switch
            {
                0 => Red,
                1 => Pink,
                2 => Grey,
                _ => Black,
            }, 0, 1),

            "3Signs" => (band switch
            {
                0 => IconShape.Diamond,
                1 => IconShape.Triangle,
                _ => IconShape.Circle,
            }, Ramp(share), 0, 1),

            "3Symbols" or "3Symbols2" => (band switch
            {
                0 => IconShape.Cross,
                1 => IconShape.Exclamation,
                _ => IconShape.Check,
            }, Ramp(share), 0, 1),

            "4Rating" or "5Rating" => (IconShape.Rating, Black, 0, (band + 1) / (double)count),

            "5Quarters" => (IconShape.Quarters, Black, 0, share),

            "3Stars" => (IconShape.Star, Amber, 0, share),

            "3Triangles" => band switch
            {
                0 => (IconShape.Triangle, Red, 180, 1),
                1 => (IconShape.Dash, Amber, 0, 1),
                _ => (IconShape.Triangle, Green, 0, 1),
            },

            "5Boxes" => (IconShape.Boxes, Blue, 0, share),

            // x14 uses this sentinel when a custom band deliberately has no icon.
            "NoIcons" => (IconShape.None, Black, 0, 0),

            _ => null,
        };
    }

    /// <summary>The colour of a band on the red-through-green ramp.</summary>
    private static PdfColor Ramp(double share) => share switch
    {
        <= 0.01 => Red,
        >= 0.99 => Green,
        < 0.5 => Blend(Red, Amber, share * 2),
        _ => Blend(Amber, Green, (share - 0.5) * 2),
    };

    /// <summary>
    /// How far an arrow is turned: down for the lowest band, up for the highest, and level
    /// through the middle, which is how Excel draws its three-, four- and five-arrow sets.
    /// </summary>
    private static double ArrowAngle(double share) => 180 - (share * 180);

    private static PdfColor Blend(PdfColor from, PdfColor to, double amount)
    {
        double mix = Math.Clamp(amount, 0, 1);
        return PdfColor.Rgb(
            (float)((from.Component0 * (1 - mix)) + (to.Component0 * mix)),
            (float)((from.Component1 * (1 - mix)) + (to.Component1 * mix)),
            (float)((from.Component2 * (1 - mix)) + (to.Component2 * mix)));
    }
}
