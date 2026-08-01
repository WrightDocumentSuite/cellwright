namespace Cellwright.Pdf.Layout;

/// <summary>
/// The units a spreadsheet measures itself in, turned into the points a page is measured in.
/// </summary>
/// <remarks>
/// A column width is not a length: it is a count of digits in the workbook's standard font, and
/// what that comes to in points depends on how wide that font draws a digit. The conversion is
/// the one the file format documents (ISO/IEC 29500-1 §18.3.1.13) — pixels first, at 96 to the
/// inch, then points at 72 — and it is why the same file prints a little differently under a
/// different standard font, in Excel as much as here.
/// </remarks>
internal static class SheetUnits
{
    /// <summary>Points per inch, which is what the margins of a page setup count in.</summary>
    public const double PointsPerInch = 72;

    /// <summary>Pixels per inch at the resolution the format assumes for widths.</summary>
    public const double PixelsPerInch = 96;

    /// <summary>The column width Excel gives a sheet that states none, in digits.</summary>
    public const double DefaultColumnWidth = 8.43;

    /// <summary>The row height Excel gives an 11-point standard font, in points.</summary>
    public const double DefaultRowHeight = 15;

    /// <summary>English Metric Units in one point, which is what anything not a cell measures in.</summary>
    public const double EmuPerPoint = 12700;

    /// <summary>
    /// How wide the standard font draws a digit when nothing has measured it: the figure for the
    /// eleven-point face Excel starts a workbook with.
    /// </summary>
    public const double NominalDigitPixels = 7;

    /// <summary>Points per pixel.</summary>
    public static double PointsPerPixel => PointsPerInch / PixelsPerInch;

    /// <summary>
    /// The width of a column, in points, from the digit count the file states.
    /// </summary>
    /// <param name="width">The width in digits of the standard font.</param>
    /// <param name="maximumDigitWidth">How wide the standard font draws a digit, in pixels.</param>
    public static double ColumnPoints(double width, double maximumDigitWidth)
    {
        double digit = Math.Max(1, maximumDigitWidth);

        // The width the format stores already has the cell padding folded into it, and rounds to
        // whole pixels: Truncate((width * digit + Truncate(128 / digit)) / 256 * 256) / digit.
        double padding = Math.Truncate(128 / digit);
        double pixels = Math.Truncate(((width * 256) + padding) / 256 * digit);
        return pixels * PointsPerPixel;
    }

    /// <summary>
    /// The width in digits a column has to state to be at least this many points wide — the
    /// inverse of <see cref="ColumnPoints"/>, which is what auto-fitting needs.
    /// </summary>
    /// <param name="points">How wide the column has to come out, in points.</param>
    /// <param name="maximumDigitWidth">How wide the standard font draws a digit, in pixels.</param>
    /// <remarks>
    /// The forward conversion truncates twice — once folding the padding in, once landing on whole
    /// pixels — so inverting it algebraically can land a hundredth short, and a column a hundredth
    /// short of its text is a column that cuts it. The analytic answer is therefore only the
    /// starting point: it is nudged up until the forward conversion agrees, which takes a step or
    /// two and makes the guarantee exact rather than approximate.
    /// </remarks>
    public static double ColumnWidth(double points, double maximumDigitWidth)
    {
        double digit = Math.Max(1, maximumDigitWidth);
        double padding = Math.Truncate(128 / digit);
        double pixels = points / PointsPerPixel;
        double width = Math.Max(0, (pixels / digit) - (padding / 256));

        width = Math.Round(width, 2, MidpointRounding.AwayFromZero);
        while (width < SheetLimits.MaxColumnWidth && ColumnPoints(width, digit) < points)
            width = Math.Round(width + 0.01, 2, MidpointRounding.AwayFromZero);

        return Math.Min(width, SheetLimits.MaxColumnWidth);
    }

    /// <summary>How wide the standard font draws a digit, in pixels at 96 per inch.</summary>
    /// <param name="digitWidthPoints">The width of a digit in points, from the font's metrics.</param>
    public static double DigitPixels(double digitWidthPoints) =>
        Math.Max(1, Math.Round(digitWidthPoints / PointsPerPixel, MidpointRounding.AwayFromZero));

    /// <summary>Turns inches — the unit the print margins count in — into points.</summary>
    /// <param name="inches">The length in inches.</param>
    public static double Points(double inches) => inches * PointsPerInch;

    /// <summary>Points per millimetre, which is what the metric papers are named in.</summary>
    public const double PointsPerMillimetre = PointsPerInch / 25.4;

    /// <summary>The paper printed on when nothing names one, and when what is named is unknown.</summary>
    public static (double Width, double Height) A4 => Millimetres(210, 297);

    /// <summary>The paper a size code names, in points, the way up the code names it.</summary>
    /// <param name="paperSize">The code the page setup states; zero means the printer's default.</param>
    /// <remarks>
    /// A code nobody here knows falls back to A4, which is what most of the world prints on.
    /// Callers that want to say so use <see cref="TryPaper"/>.
    /// </remarks>
    public static (double Width, double Height) Paper(int paperSize) => Named(paperSize) ?? A4;

    /// <summary>
    /// The paper a size code names (ISO/IEC 29500-1 §18.3.1.63), and whether the code is one the
    /// format names at all.
    /// </summary>
    /// <param name="paperSize">The code the page setup states; zero means the printer's default.</param>
    /// <param name="size">The paper, in points, or A4 when the code is not one of the list.</param>
    /// <remarks>
    /// A few of the codes are named in the list without a measurement — the Japanese envelopes —
    /// and those carry the sizes their names have always stood for. The rotated codes are the
    /// same papers the other way up, and they are kept that way here rather than turned back:
    /// a sheet asking for A4 rotated is asking for a landscape page, whatever its orientation
    /// says.
    /// </remarks>
    public static bool TryPaper(int paperSize, out (double Width, double Height) size)
    {
        (double Width, double Height)? named = Named(paperSize);
        size = named ?? A4;
        return named is not null;
    }

    private static (double Width, double Height)? Named(int code) => code switch
    {
        // Zero is the printer's own default, which is nothing this side of a printer can know.
        0 => A4,

        1 or 2 or 18 => Inches(8.5, 11),        // Letter, Letter small, Note
        3 => Inches(11, 17),                    // Tabloid
        4 => Inches(17, 11),                    // Ledger, which is Tabloid on its side
        5 => Inches(8.5, 14),                   // Legal
        6 => Inches(5.5, 8.5),                  // Statement
        7 => Inches(7.25, 10.5),                // Executive
        8 => Millimetres(297, 420),             // A3
        9 or 10 => Millimetres(210, 297),       // A4, A4 small
        11 => Millimetres(148, 210),            // A5
        12 => Millimetres(250, 353),            // B4
        13 => Millimetres(176, 250),            // B5
        14 => Inches(8.5, 13),                  // Folio
        15 => Millimetres(215, 275),            // Quarto
        16 => Inches(10, 14),                   // Standard
        17 => Inches(11, 17),                   // Standard
        19 => Inches(3.875, 8.875),             // #9 envelope
        20 => Inches(4.125, 9.5),               // #10 envelope
        21 => Inches(4.5, 10.375),              // #11 envelope
        22 => Inches(4.75, 11),                 // #12 envelope
        23 => Inches(5, 11.5),                  // #14 envelope
        24 => Inches(17, 22),                   // C
        25 => Inches(22, 34),                   // D
        26 => Inches(34, 44),                   // E
        27 => Millimetres(110, 220),            // DL envelope
        28 => Millimetres(162, 229),            // C5 envelope
        29 => Millimetres(324, 458),            // C3 envelope
        30 => Millimetres(229, 324),            // C4 envelope
        31 => Millimetres(114, 162),            // C6 envelope
        32 => Millimetres(114, 229),            // C65 envelope
        33 or 42 => Millimetres(250, 353),      // B4 envelope, ISO B4
        34 => Millimetres(176, 250),            // B5 envelope
        35 => Millimetres(176, 125),            // B6 envelope
        36 => Millimetres(110, 230),            // Italy envelope
        37 => Inches(3.875, 7.5),               // Monarch envelope
        38 => Inches(3.625, 6.5),               // 6 3/4 envelope
        39 => Inches(14.875, 11),               // US standard fanfold
        40 => Inches(8.5, 12),                  // German standard fanfold
        41 => Inches(8.5, 13),                  // German legal fanfold
        43 or 69 => Millimetres(200, 148),      // Japanese double postcard
        44 => Inches(9, 11),                    // Standard
        45 => Inches(10, 11),                   // Standard
        46 => Inches(15, 11),                   // Standard
        47 => Millimetres(220, 220),            // Invite envelope
        50 or 56 => Inches(9.275, 12),          // Letter extra, Letter extra transverse
        51 => Inches(9.275, 15),                // Legal extra
        52 => Inches(11.69, 18),                // Tabloid extra
        53 => Millimetres(236, 322),            // A4 extra
        54 => Inches(8.275, 11),                // Letter transverse
        55 => Millimetres(210, 297),            // A4 transverse
        57 => Millimetres(227, 356),            // SuperA/SuperA/A4
        58 => Millimetres(305, 487),            // SuperB/SuperB/A3
        59 => Inches(8.5, 12.69),               // Letter plus
        60 => Millimetres(210, 330),            // A4 plus
        61 => Millimetres(148, 210),            // A5 transverse
        62 => Millimetres(182, 257),            // JIS B5 transverse
        63 or 68 => Millimetres(322, 445),      // A3 extra, A3 extra transverse
        64 => Millimetres(174, 235),            // A5 extra
        65 => Millimetres(201, 276),            // ISO B5 extra
        66 => Millimetres(420, 594),            // A2
        67 => Millimetres(297, 420),            // A3 transverse
        70 => Millimetres(105, 148),            // A6
        71 => Millimetres(240, 332),            // Japanese envelope Kaku #2
        72 => Millimetres(216, 277),            // Japanese envelope Kaku #3
        73 => Millimetres(120, 235),            // Japanese envelope Chou #3
        74 => Millimetres(90, 205),             // Japanese envelope Chou #4
        75 => Inches(11, 8.5),                  // Letter rotated
        76 => Millimetres(420, 297),            // A3 rotated
        77 => Millimetres(297, 210),            // A4 rotated
        78 => Millimetres(210, 148),            // A5 rotated
        79 => Millimetres(364, 257),            // B4 (JIS) rotated
        80 => Millimetres(257, 182),            // B5 (JIS) rotated
        81 => Millimetres(148, 100),            // Japanese postcard rotated
        82 => Millimetres(148, 200),            // Double Japanese postcard rotated
        83 => Millimetres(148, 105),            // A6 rotated
        84 => Millimetres(332, 240),            // Japanese envelope Kaku #2 rotated
        85 => Millimetres(277, 216),            // Japanese envelope Kaku #3 rotated
        86 => Millimetres(235, 120),            // Japanese envelope Chou #3 rotated
        87 => Millimetres(205, 90),             // Japanese envelope Chou #4 rotated
        88 => Millimetres(128, 182),            // B6 (JIS)
        89 => Millimetres(182, 128),            // B6 (JIS) rotated
        90 => Inches(12, 11),                   // Standard
        91 => Millimetres(105, 235),            // Japanese envelope You #4
        92 => Millimetres(235, 105),            // Japanese envelope You #4 rotated
        93 => Millimetres(146, 215),            // PRC 16K
        94 or 95 => Millimetres(97, 151),       // PRC 32K, PRC 32K (Big)
        96 => Millimetres(102, 165),            // PRC envelope #1
        97 => Millimetres(102, 176),            // PRC envelope #2
        98 => Millimetres(125, 176),            // PRC envelope #3
        99 => Millimetres(110, 208),            // PRC envelope #4
        100 => Millimetres(110, 220),           // PRC envelope #5
        101 => Millimetres(120, 230),           // PRC envelope #6
        102 => Millimetres(160, 230),           // PRC envelope #7
        103 => Millimetres(120, 309),           // PRC envelope #8
        104 => Millimetres(229, 324),           // PRC envelope #9
        105 => Millimetres(324, 458),           // PRC envelope #10
        106 => Millimetres(215, 146),           // PRC 16K rotated
        107 or 108 => Millimetres(151, 97),     // PRC 32K rotated, PRC 32K (Big) rotated
        109 => Millimetres(165, 102),           // PRC envelope #1 rotated
        110 => Millimetres(176, 102),           // PRC envelope #2 rotated
        111 => Millimetres(176, 125),           // PRC envelope #3 rotated
        112 => Millimetres(208, 110),           // PRC envelope #4 rotated
        113 => Millimetres(220, 110),           // PRC envelope #5 rotated
        114 => Millimetres(230, 120),           // PRC envelope #6 rotated
        115 => Millimetres(230, 160),           // PRC envelope #7 rotated
        116 => Millimetres(309, 120),           // PRC envelope #8 rotated
        117 => Millimetres(324, 229),           // PRC envelope #9 rotated
        118 => Millimetres(458, 324),           // PRC envelope #10 rotated
        _ => null,
    };

    /// <summary>A paper named in inches, in points.</summary>
    private static (double Width, double Height) Inches(double width, double height) =>
        (width * PointsPerInch, height * PointsPerInch);

    /// <summary>
    /// A paper named in millimetres, in whole points — which is how every reader states A4.
    /// </summary>
    private static (double Width, double Height) Millimetres(double width, double height) =>
        (Math.Round(width * PointsPerMillimetre), Math.Round(height * PointsPerMillimetre));
}
