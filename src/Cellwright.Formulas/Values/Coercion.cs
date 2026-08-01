using System.Globalization;

namespace Cellwright.Formulas;

/// <summary>
/// Excel's conversion rules between cell values. These are the subtle parts that make an
/// engine agree with Excel: blanks behave as zero or empty text depending on context, text
/// that looks numeric converts, and errors travel outwards untouched.
/// </summary>
internal static class Coercion
{
    /// <summary>Converts to a number, or returns an error value.</summary>
    public static CellValue ToNumber(in CellValue value, ExcelDateSystem dateSystem) => value.Kind switch
    {
        CellValueKind.Number => value,
        CellValueKind.Blank => CellValue.Number(0),
        CellValueKind.Boolean => CellValue.Number(value.GetBoolean() ? 1 : 0),
        CellValueKind.DateTime or CellValueKind.TimeSpan => CellValue.Number(value.ToSerialNumber(dateSystem)),
        CellValueKind.Error => value,
        _ => ParseNumber(value.GetText()),
    };

    /// <summary>Converts to text the way <c>&amp;</c> and text functions do.</summary>
    public static string ToText(in CellValue value, ExcelDateSystem dateSystem) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Text => value.GetText(),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Number => FormatNumber(value.GetNumber()),
        CellValueKind.DateTime or CellValueKind.TimeSpan => FormatNumber(value.ToSerialNumber(dateSystem)),
        _ => value.ToString(),
    };

    /// <summary>Converts to a logical value, or returns an error.</summary>
    public static CellValue ToBoolean(in CellValue value) => value.Kind switch
    {
        CellValueKind.Boolean => value,
        CellValueKind.Blank => CellValue.Boolean(false),
        CellValueKind.Number => CellValue.Boolean(value.GetNumber() != 0),
        CellValueKind.Error => value,
        CellValueKind.Text => ParseBoolean(value.GetText()),
        _ => CellValue.Boolean(value.ToSerialNumber() != 0),
    };

    /// <summary>
    /// Compares two values with Excel's ordering: numbers sort before text, which sorts before
    /// logicals. A blank takes the type of the value it is compared with.
    /// </summary>
    public static int Compare(in CellValue left, in CellValue right, ExcelDateSystem dateSystem)
    {
        int leftRank = TypeRank(left, right);
        int rightRank = TypeRank(right, left);
        if (leftRank != rightRank)
            return leftRank.CompareTo(rightRank);

        return leftRank switch
        {
            0 => ToNumberOrZero(left, dateSystem).CompareTo(ToNumberOrZero(right, dateSystem)),
            1 => string.Compare(TextOrEmpty(left), TextOrEmpty(right), StringComparison.OrdinalIgnoreCase),
            _ => BooleanOrFalse(left).CompareTo(BooleanOrFalse(right)),
        };
    }

    /// <summary>Whether the value counts as text for comparison purposes.</summary>
    private static int TypeRank(in CellValue value, in CellValue other) => value.Kind switch
    {
        CellValueKind.Text => 1,
        CellValueKind.Boolean => 2,
        CellValueKind.Blank => other.Kind switch
        {
            CellValueKind.Text => 1,
            CellValueKind.Boolean => 2,
            _ => 0,
        },
        _ => 0,
    };

    private static double ToNumberOrZero(in CellValue value, ExcelDateSystem dateSystem)
    {
        CellValue number = ToNumber(value, dateSystem);
        return number.IsNumber ? number.GetNumber() : 0;
    }

    private static string TextOrEmpty(in CellValue value) => value.IsText ? value.GetText() : "";

    private static bool BooleanOrFalse(in CellValue value) => value.IsBoolean && value.GetBoolean();

    /// <summary>
    /// Parses text the way arithmetic does. Excel accepts surrounding spaces, a percent suffix
    /// and thousands separators; anything else is a <c>#VALUE!</c>.
    /// </summary>
    public static CellValue ParseNumber(string text)
    {
        ReadOnlySpan<char> span = text.AsSpan().Trim();
        if (span.IsEmpty)
            return CellValue.Number(0);

        bool percent = span[^1] == '%';
        if (percent)
            span = span[..^1].TrimEnd();

        const NumberStyles Styles = NumberStyles.Float | NumberStyles.AllowThousands;
        if (!double.TryParse(span, Styles, CultureInfo.InvariantCulture, out double number))
            return CellValue.Error(ExcelError.Value);

        return CellValue.Number(percent ? number / 100 : number);
    }

    private static CellValue ParseBoolean(string text) => text.Trim() switch
    {
        var s when s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) => CellValue.Boolean(true),
        var s when s.Equals("FALSE", StringComparison.OrdinalIgnoreCase) => CellValue.Boolean(false),
        _ => CellValue.Error(ExcelError.Value),
    };

    /// <summary>
    /// Renders a number the way Excel's General format does for concatenation: up to fifteen
    /// significant digits, no trailing zeros, scientific notation only at the extremes.
    /// </summary>
    public static string FormatNumber(double value)
    {
        if (value == 0)
            return "0";

        double magnitude = Math.Abs(value);
        if (magnitude >= 1e11 || magnitude < 1e-10)
            return value.ToString("0.#####E+00", CultureInfo.InvariantCulture);

        return value.ToString("G15", CultureInfo.InvariantCulture);
    }
}
