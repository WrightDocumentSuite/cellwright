namespace Cellwright.Mapping;

/// <summary>
/// Conversion helpers the generated schemas call. Centralizing them keeps generated code
/// small and error messages uniform. Blank cells convert to <see langword="null"/> for
/// nullable targets and throw for non-nullable ones.
/// </summary>
public static class RowConvert
{
    /// <summary>Converts to <see cref="string"/>; blank becomes an empty string, other kinds use their display text.</summary>
    public static string ToText(in CellValue value, string column) => value.ToString();

    /// <summary>Converts to <see cref="string"/> or <see langword="null"/> for blank.</summary>
    public static string? ToTextOrNull(in CellValue value, string column) => value.IsBlank ? null : value.ToString();

    /// <summary>Converts to <see cref="double"/>.</summary>
    public static double ToDouble(in CellValue value, string column) =>
        value.TryGetNumber(out double number) ? number : throw Mismatch(value, column, "a number");

    /// <summary>Converts to <see cref="double"/> or <see langword="null"/> for blank.</summary>
    public static double? ToDoubleOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToDouble(value, column);

    /// <summary>Converts to <see cref="int"/>.</summary>
    public static int ToInt32(in CellValue value, string column) => checked((int)Math.Round(ToDouble(value, column)));

    /// <summary>Converts to <see cref="int"/> or <see langword="null"/> for blank.</summary>
    public static int? ToInt32OrNull(in CellValue value, string column) => value.IsBlank ? null : ToInt32(value, column);

    /// <summary>Converts to <see cref="long"/>.</summary>
    public static long ToInt64(in CellValue value, string column) => checked((long)Math.Round(ToDouble(value, column)));

    /// <summary>Converts to <see cref="long"/> or <see langword="null"/> for blank.</summary>
    public static long? ToInt64OrNull(in CellValue value, string column) => value.IsBlank ? null : ToInt64(value, column);

    /// <summary>Converts to <see cref="decimal"/>.</summary>
    public static decimal ToDecimal(in CellValue value, string column) => (decimal)ToDouble(value, column);

    /// <summary>Converts to <see cref="decimal"/> or <see langword="null"/> for blank.</summary>
    public static decimal? ToDecimalOrNull(in CellValue value, string column) => value.IsBlank ? null : ToDecimal(value, column);

    /// <summary>Converts to <see cref="bool"/>.</summary>
    public static bool ToBoolean(in CellValue value, string column) =>
        value.IsBoolean ? value.GetBoolean() : throw Mismatch(value, column, "a boolean");

    /// <summary>Converts to <see cref="bool"/> or <see langword="null"/> for blank.</summary>
    public static bool? ToBooleanOrNull(in CellValue value, string column) => value.IsBlank ? null : ToBoolean(value, column);

    /// <summary>Converts to <see cref="DateTime"/>.</summary>
    public static DateTime ToDateTime(in CellValue value, string column) =>
        value.TryGetDateTime(out DateTime timestamp) ? timestamp : throw Mismatch(value, column, "a date");

    /// <summary>Converts to <see cref="DateTime"/> or <see langword="null"/> for blank.</summary>
    public static DateTime? ToDateTimeOrNull(in CellValue value, string column) => value.IsBlank ? null : ToDateTime(value, column);

    /// <summary>Converts to <see cref="DateOnly"/>.</summary>
    public static DateOnly ToDateOnly(in CellValue value, string column) => DateOnly.FromDateTime(ToDateTime(value, column));

    /// <summary>Converts to <see cref="DateOnly"/> or <see langword="null"/> for blank.</summary>
    public static DateOnly? ToDateOnlyOrNull(in CellValue value, string column) => value.IsBlank ? null : ToDateOnly(value, column);

    /// <summary>Converts to <see cref="TimeSpan"/>.</summary>
    public static TimeSpan ToTimeSpan(in CellValue value, string column) =>
        value.IsTimeSpan ? value.GetTimeSpan() : throw Mismatch(value, column, "a duration");

    /// <summary>Converts to <see cref="TimeSpan"/> or <see langword="null"/> for blank.</summary>
    public static TimeSpan? ToTimeSpanOrNull(in CellValue value, string column) => value.IsBlank ? null : ToTimeSpan(value, column);

    /// <summary>Passes the raw cell value through.</summary>
    public static CellValue ToCellValue(in CellValue value, string column) => value;

    private static CellMappingException Mismatch(in CellValue value, string column, string expected) =>
        new($"Column '{column}': cannot convert {value.Kind} value '{value}' to {expected}.");
}
