namespace Cellwright;

/// <summary>
/// The colour scheme of a workbook's theme, which a themed <see cref="ExcelColor"/> names a slot
/// of (ECMA-376 Part 1 §20.1.6.2).
/// </summary>
/// <remarks>
/// A theme colour is not an exotic case: it is what nearly every styled file uses, because
/// Excel's own palette is built on one. Without the scheme a slot number is all a reader has,
/// and two workbooks under different themes give the same number to different colours.
/// <para>
/// The order the slots are numbered in is not the order the file states them. The scheme opens
/// with the two dark colours and the two light ones, but a style refers to them with the pairs
/// swapped — index 0 is the light background, index 1 the dark text. The swap is in the format,
/// not in this reading of it: §18.8.3 numbers the styles' view, §20.1.6.2 the theme's.
/// </para>
/// </remarks>
public sealed class WorkbookTheme
{
    /// <summary>Number of slots a scheme has.</summary>
    public const int SlotCount = 12;

    private readonly ExcelColor[] _slots;

    private WorkbookTheme(ExcelColor[] slots) => _slots = slots;

    /// <summary>The theme Excel starts a new workbook with.</summary>
    public static WorkbookTheme Office { get; } = new(
    [
        ExcelColor.FromArgb(0xFFFFFFFF),        // 0 background 1 (lt1)
        ExcelColor.FromArgb(0xFF000000),        // 1 text 1 (dk1)
        ExcelColor.FromArgb(0xFFEEECE1),        // 2 background 2 (lt2)
        ExcelColor.FromArgb(0xFF1F497D),        // 3 text 2 (dk2)
        ExcelColor.FromArgb(0xFF4F81BD),        // 4 accent 1
        ExcelColor.FromArgb(0xFFC0504D),        // 5 accent 2
        ExcelColor.FromArgb(0xFF9BBB59),        // 6 accent 3
        ExcelColor.FromArgb(0xFF8064A2),        // 7 accent 4
        ExcelColor.FromArgb(0xFF4BACC6),        // 8 accent 5
        ExcelColor.FromArgb(0xFFF79646),        // 9 accent 6
        ExcelColor.FromArgb(0xFF0000FF),        // 10 hyperlink
        ExcelColor.FromArgb(0xFF800080),        // 11 followed hyperlink
    ]);

    /// <summary>The colour of a slot, by the index a style names it with.</summary>
    public ExcelColor this[int index] =>
        (uint)index < (uint)_slots.Length ? _slots[index] : ExcelColor.Auto;

    /// <summary>
    /// Resolves a colour to an explicit one. An RGB colour is returned as it is, an automatic
    /// one stays automatic, and a theme slot becomes the scheme's colour with its tint applied.
    /// </summary>
    public ExcelColor Resolve(ExcelColor color)
    {
        if (!color.IsThemed)
            return color;

        ExcelColor slot = this[color.ThemeIndex];
        return slot.IsAuto || color.Tint == 0 ? slot : ExcelColor.FromArgb(ApplyTint(slot.Argb, color.Tint));
    }

    /// <summary>
    /// Lightens or darkens a colour by a tint, as ECMA-376 Part 1 §18.8.3 defines it: the tint
    /// moves the luminance of the colour in HSL space, towards white when positive and towards
    /// black when negative, leaving hue and saturation alone.
    /// </summary>
    public static uint ApplyTint(uint argb, double tint)
    {
        if (tint == 0)
            return argb;

        tint = Math.Clamp(tint, -1, 1);
        (double hue, double saturation, double luminance) = ToHsl(argb);

        // Negative shrinks the luminance towards black; positive moves what is left of the way
        // to white. Both are the spec's formulas with the maximum taken as one rather than 255.
        luminance = tint < 0 ? luminance * (1 + tint) : (luminance * (1 - tint)) + tint;

        return (argb & 0xFF000000u) | FromHsl(hue, saturation, Math.Clamp(luminance, 0, 1));
    }

    /// <summary>
    /// Reads the colour scheme of a theme part. Anything missing keeps the Office default, so a
    /// theme that states only its accents still answers for the rest.
    /// </summary>
    public static WorkbookTheme Read(Stream stream)
    {
        ExcelColor[] slots = [.. Office._slots];
        using var xml = System.Xml.XmlReader.Create(stream, Xml.XmlDefaults.ReaderSettings);

        // Only the scheme matters here, and a theme part also carries fonts and format schemes
        // holding colours of their own — so reading stops at the end of clrScheme.
        while (xml.Read())
        {
            if (xml.NodeType == System.Xml.XmlNodeType.Element && xml.LocalName == "clrScheme")
            {
                ReadScheme(xml, slots);
                break;
            }
        }

        return new WorkbookTheme(slots);
    }

    private static void ReadScheme(System.Xml.XmlReader xml, ExcelColor[] slots)
    {
        int depth = xml.Depth;
        int? pending = null;

        while (xml.Read() && xml.Depth > depth)
        {
            if (xml.NodeType != System.Xml.XmlNodeType.Element)
                continue;

            if (SlotOf(xml.LocalName) is { } slot)
            {
                pending = slot;
                continue;
            }

            if (pending is not { } target)
                continue;

            // A system colour states what it resolved to last time it was rendered, which is
            // the only value a file can offer for something the operating system decides.
            string? hex = xml.LocalName switch
            {
                "srgbClr" => xml.GetAttribute("val"),
                "sysClr" => xml.GetAttribute("lastClr"),
                _ => null,
            };

            if (hex is not null && ExcelColor.TryParseHex(hex, out ExcelColor parsed))
                slots[target] = parsed;
            pending = null;
        }
    }

    /// <summary>
    /// The slot index a scheme element fills. This is where the dark and light pairs swap: the
    /// theme states dk1 first, but a style calls the light one index 0.
    /// </summary>
    private static int? SlotOf(string name) => name switch
    {
        "lt1" => 0,
        "dk1" => 1,
        "lt2" => 2,
        "dk2" => 3,
        "accent1" => 4,
        "accent2" => 5,
        "accent3" => 6,
        "accent4" => 7,
        "accent5" => 8,
        "accent6" => 9,
        "hlink" => 10,
        "folHlink" => 11,
        _ => null,
    };

    private static (double Hue, double Saturation, double Luminance) ToHsl(uint argb)
    {
        double red = ((argb >> 16) & 0xFF) / 255.0;
        double green = ((argb >> 8) & 0xFF) / 255.0;
        double blue = (argb & 0xFF) / 255.0;

        double max = Math.Max(red, Math.Max(green, blue));
        double min = Math.Min(red, Math.Min(green, blue));
        double luminance = (max + min) / 2;
        if (max == min)
            return (0, 0, luminance);

        double delta = max - min;
        double saturation = luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min);
        double hue;
        if (max == red)
            hue = ((green - blue) / delta) + (green < blue ? 6 : 0);
        else if (max == green)
            hue = ((blue - red) / delta) + 2;
        else
            hue = ((red - green) / delta) + 4;

        return (hue / 6, saturation, luminance);
    }

    private static uint FromHsl(double hue, double saturation, double luminance)
    {
        if (saturation == 0)
        {
            byte grey = Channel(luminance);
            return (uint)((grey << 16) | (grey << 8) | grey);
        }

        double q = luminance < 0.5 ? luminance * (1 + saturation) : luminance + saturation - (luminance * saturation);
        double p = (2 * luminance) - q;
        return (uint)((Channel(Component(p, q, hue + (1.0 / 3))) << 16)
            | (Channel(Component(p, q, hue)) << 8)
            | Channel(Component(p, q, hue - (1.0 / 3))));
    }

    private static double Component(double p, double q, double t)
    {
        if (t < 0)
            t += 1;
        if (t > 1)
            t -= 1;
        if (t < 1.0 / 6)
            return p + ((q - p) * 6 * t);
        if (t < 1.0 / 2)
            return q;
        if (t < 2.0 / 3)
            return p + ((q - p) * ((2.0 / 3) - t) * 6);
        return p;
    }

    private static byte Channel(double value) => (byte)Math.Clamp(Math.Round(value * 255), 0, 255);
}
