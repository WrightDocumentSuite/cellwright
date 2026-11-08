using System.Globalization;

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

    /// <summary>Converts to <see cref="float"/>.</summary>
    public static float ToSingle(in CellValue value, string column) => (float)ToDouble(value, column);

    /// <summary>Converts to <see cref="float"/> or <see langword="null"/> for blank.</summary>
    public static float? ToSingleOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToSingle(value, column);

    /// <summary>Converts to <see cref="short"/>.</summary>
    public static short ToInt16(in CellValue value, string column) => checked((short)Math.Round(ToDouble(value, column)));

    /// <summary>Converts to <see cref="short"/> or <see langword="null"/> for blank.</summary>
    public static short? ToInt16OrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToInt16(value, column);

    /// <summary>Converts to <see cref="byte"/>.</summary>
    public static byte ToByte(in CellValue value, string column) => checked((byte)Math.Round(ToDouble(value, column)));

    /// <summary>Converts to <see cref="byte"/> or <see langword="null"/> for blank.</summary>
    public static byte? ToByteOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToByte(value, column);

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

    /// <summary>Converts to <see cref="DateTimeOffset"/>.</summary>
    public static DateTimeOffset ToDateTimeOffset(in CellValue value, string column)
    {
        if (value.TryGetDateTime(out DateTime timestamp))
            return new DateTimeOffset(DateTime.SpecifyKind(timestamp, DateTimeKind.Utc));
        if (value.TryGetText(out string text) &&
            DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset parsed))
        {
            return parsed;
        }

        throw Mismatch(value, column, "a date-time offset");
    }

    /// <summary>Converts to <see cref="DateTimeOffset"/> or <see langword="null"/> for blank.</summary>
    public static DateTimeOffset? ToDateTimeOffsetOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToDateTimeOffset(value, column);

    /// <summary>Converts to <see cref="DateOnly"/>.</summary>
    public static DateOnly ToDateOnly(in CellValue value, string column) => DateOnly.FromDateTime(ToDateTime(value, column));

    /// <summary>Converts to <see cref="DateOnly"/> or <see langword="null"/> for blank.</summary>
    public static DateOnly? ToDateOnlyOrNull(in CellValue value, string column) => value.IsBlank ? null : ToDateOnly(value, column);

    /// <summary>Converts to <see cref="TimeOnly"/>.</summary>
    public static TimeOnly ToTimeOnly(in CellValue value, string column)
    {
        if (value.IsTimeSpan)
            return TimeOnly.FromTimeSpan(value.GetTimeSpan());
        if (value.TryGetDateTime(out DateTime timestamp))
            return TimeOnly.FromDateTime(timestamp);
        throw Mismatch(value, column, "a time of day");
    }

    /// <summary>Converts to <see cref="TimeOnly"/> or <see langword="null"/> for blank.</summary>
    public static TimeOnly? ToTimeOnlyOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToTimeOnly(value, column);

    /// <summary>Converts to <see cref="TimeSpan"/>.</summary>
    public static TimeSpan ToTimeSpan(in CellValue value, string column) =>
        value.IsTimeSpan ? value.GetTimeSpan() : throw Mismatch(value, column, "a duration");

    /// <summary>Converts to <see cref="TimeSpan"/> or <see langword="null"/> for blank.</summary>
    public static TimeSpan? ToTimeSpanOrNull(in CellValue value, string column) => value.IsBlank ? null : ToTimeSpan(value, column);

    /// <summary>Converts to <see cref="Guid"/>.</summary>
    public static Guid ToGuid(in CellValue value, string column) =>
        value.TryGetText(out string text) && Guid.TryParse(text, out Guid guid)
            ? guid
            : throw Mismatch(value, column, "a GUID");

    /// <summary>Converts to <see cref="Guid"/> or <see langword="null"/> for blank.</summary>
    public static Guid? ToGuidOrNull(in CellValue value, string column) =>
        value.IsBlank ? null : ToGuid(value, column);

    /// <summary>Converts to <typeparamref name="TEnum"/> by member name (case-insensitive) or integral value.</summary>
    public static TEnum ToEnum<TEnum>(in CellValue value, string column)
        where TEnum : struct, Enum
    {
        if (value.TryGetText(out string text) && Enum.TryParse(text, ignoreCase: true, out TEnum parsed))
            return parsed;
        if (value.TryGetNumber(out double number))
        {
            long integral = checked((long)Math.Round(number));
            return (TEnum)Enum.ToObject(typeof(TEnum), integral);
        }

        throw Mismatch(value, column, $"an enum of type {typeof(TEnum).Name}");
    }

    /// <summary>Converts to <typeparamref name="TEnum"/> or <see langword="null"/> for blank.</summary>
    public static TEnum? ToEnumOrNull<TEnum>(in CellValue value, string column)
        where TEnum : struct, Enum =>
        value.IsBlank ? null : ToEnum<TEnum>(value, column);

    /// <summary>Passes the raw cell value through.</summary>
    public static CellValue ToCellValue(in CellValue value, string column) => value;

    private static CellMappingException Mismatch(in CellValue value, string column, string expected) =>
        new($"Column '{column}': cannot convert {value.Kind} value '{value}' to {expected}.");
}
