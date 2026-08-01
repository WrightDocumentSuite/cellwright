using System.Globalization;

namespace Cellwright;

/// <summary>How a color says what it is.</summary>
public enum ExcelColorKind : byte
{
    /// <summary>Automatic: the consumer picks, normally black text on a white sheet.</summary>
    Auto,

    /// <summary>An explicit ARGB value.</summary>
    Rgb,

    /// <summary>A slot in the document theme, optionally lightened or darkened.</summary>
    Theme,
}

/// <summary>
/// A color in a workbook: automatic, an explicit ARGB value, or a theme slot with a tint.
/// </summary>
/// <remarks>
/// Excel's own palette is theme-based, so a theme color is not an exotic case — it is what
/// nearly every styled file uses. The theme part is carried through the package untouched, so a
/// color keeps its slot and comes back looking the same even though nothing here resolves it to
/// pixels. The legacy <c>indexed</c> form is resolved on the way in, against the fixed palette
/// every producer of such a file agreed on.
/// </remarks>
public readonly record struct ExcelColor
{
    private readonly uint _argb;
    private readonly int _theme;
    private readonly double _tint;
    private readonly ExcelColorKind _kind;

    private ExcelColor(ExcelColorKind kind, uint argb, int theme, double tint)
    {
        _kind = kind;
        _argb = argb;
        _theme = theme;
        _tint = tint;
    }

    /// <summary>The automatic color (default text/border color decided by the consumer).</summary>
    public static ExcelColor Auto => default;

    /// <summary>How the color says what it is.</summary>
    public ExcelColorKind Kind => _kind;

    /// <summary>Whether the color is automatic rather than stated.</summary>
    public bool IsAuto => _kind == ExcelColorKind.Auto;

    /// <summary>Whether the color names a theme slot.</summary>
    public bool IsThemed => _kind == ExcelColorKind.Theme;

    /// <summary>The ARGB value of an explicit color, e.g. <c>0xFF2278D4</c>; zero for the others.</summary>
    public uint Argb => _argb;

    /// <summary>Zero-based theme slot, meaningful only when <see cref="IsThemed"/>.</summary>
    public int ThemeIndex => _theme;

    /// <summary>Lightening (positive) or darkening (negative) of a theme slot, from -1 to 1.</summary>
    public double Tint => _tint;

    /// <summary>Creates a color from a 32-bit ARGB value, e.g. <c>0xFF2278D4</c>.</summary>
    public static ExcelColor FromArgb(uint argb) => new(ExcelColorKind.Rgb, argb, 0, 0);

    /// <summary>Creates an opaque color from RGB components.</summary>
    public static ExcelColor FromRgb(byte red, byte green, byte blue) =>
        FromArgb(0xFF000000u | (uint)(red << 16) | (uint)(green << 8) | blue);

    /// <summary>Creates a color naming a theme slot, optionally lightened or darkened.</summary>
    /// <param name="index">Zero-based theme slot.</param>
    /// <param name="tint">Lightening (positive) or darkening (negative), from -1 to 1.</param>
    public static ExcelColor FromTheme(int index, double tint = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        return new ExcelColor(ExcelColorKind.Theme, 0, index, tint);
    }

    /// <summary>
    /// Resolves one of the 64 legacy palette slots to its color. Indexes 64 and 65 mean the
    /// system foreground and background, which have no fixed value and stay automatic.
    /// </summary>
    public static ExcelColor FromIndexed(int index) =>
        (uint)index < (uint)IndexedPalette.Length ? FromArgb(IndexedPalette[index]) : Auto;

    /// <summary>Formats as the 8-digit hex form used in style XML, e.g. <c>FF2278D4</c>.</summary>
    public string ToHex() => _argb.ToString("X8", CultureInfo.InvariantCulture);

    /// <summary>Parses the 8- or 6-digit hex form found in style XML.</summary>
    public static bool TryParseHex(ReadOnlySpan<char> hex, out ExcelColor color)
    {
        color = Auto;
        if (hex.Length is not (6 or 8) || !uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint value))
            return false;

        if (hex.Length == 6)
            value |= 0xFF000000u;
        color = FromArgb(value);
        return true;
    }

    /// <inheritdoc/>
    public override string ToString() => _kind switch
    {
        ExcelColorKind.Rgb => "#" + ToHex(),
        ExcelColorKind.Theme => _tint == 0
            ? $"Theme {_theme}"
            : $"Theme {_theme} {_tint.ToString("+0.##;-0.##", CultureInfo.InvariantCulture)}",
        _ => "Auto",
    };

    /// <summary>
    /// The palette an <c>indexed</c> color points into. Its first sixteen entries repeat in a
    /// different order further down — an artefact of the 1980s file format that every producer
    /// still writes, so the table is reproduced as specified rather than tidied up.
    /// </summary>
    private static ReadOnlySpan<uint> IndexedPalette =>
    [
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF800000, 0xFF008000, 0xFF000080, 0xFF808000, 0xFF800080, 0xFF008080, 0xFFC0C0C0, 0xFF808080,
        0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF, 0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF,
        0xFF000080, 0xFFFF00FF, 0xFFFFFF00, 0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF,
        0xFF00CCFF, 0xFFCCFFFF, 0xFFCCFFCC, 0xFFFFFF99, 0xFF99CCFF, 0xFFFF99CC, 0xFFCC99FF, 0xFFFFCC99,
        0xFF3366FF, 0xFF33CCCC, 0xFF99CC00, 0xFFFFCC00, 0xFFFF9900, 0xFFFF6600, 0xFF666699, 0xFF969696,
        0xFF003366, 0xFF339966, 0xFF003300, 0xFF333300, 0xFF993300, 0xFF993366, 0xFF333399, 0xFF333333,
    ];
}
