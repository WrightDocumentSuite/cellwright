namespace Cellwright;

public readonly partial struct CellValue
{
    /// <summary>Wraps a number.</summary>
    public static implicit operator CellValue(double value) => Number(value);

    /// <summary>Wraps a number.</summary>
    public static implicit operator CellValue(int value) => Number(value);

    /// <summary>Wraps a number. Values beyond 2^53 lose precision, as they do in Excel itself.</summary>
    public static implicit operator CellValue(long value) => Number(value);

    /// <summary>Wraps a number. Excel stores IEEE doubles, so extreme decimal precision is not preserved.</summary>
    public static implicit operator CellValue(decimal value) => Number((double)value);

    /// <summary>Wraps text; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(string? value) => value is null ? Blank : Text(value);

    /// <summary>Wraps a logical value.</summary>
    public static implicit operator CellValue(bool value) => Boolean(value);

    /// <summary>Wraps a date-time.</summary>
    public static implicit operator CellValue(DateTime value) => DateTime(value);

    /// <summary>Wraps a date as midnight of that day.</summary>
    public static implicit operator CellValue(DateOnly value) => DateTime(value.ToDateTime(System.TimeOnly.MinValue));

    /// <summary>Wraps a time of day as a duration since midnight.</summary>
    public static implicit operator CellValue(TimeOnly value) => TimeSpan(value.ToTimeSpan());

    /// <summary>Wraps a duration.</summary>
    public static implicit operator CellValue(TimeSpan value) => TimeSpan(value);

    /// <summary>Wraps an error.</summary>
    public static implicit operator CellValue(ExcelError value) => Error(value);

    /// <summary>Wraps a nullable number; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(double? value) => value is null ? Blank : Number(value.Value);

    /// <summary>Wraps a nullable number; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(int? value) => value is null ? Blank : Number(value.Value);

    /// <summary>Wraps a nullable number; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(long? value) => value is null ? Blank : Number(value.Value);

    /// <summary>Wraps a nullable number; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(decimal? value) => value is null ? Blank : Number((double)value.Value);

    /// <summary>Wraps a nullable logical value; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(bool? value) => value is null ? Blank : Boolean(value.Value);

    /// <summary>Wraps a nullable date-time; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(DateTime? value) => value is null ? Blank : DateTime(value.Value);

    /// <summary>Wraps a nullable duration; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(TimeSpan? value) => value is null ? Blank : TimeSpan(value.Value);

    /// <summary>Wraps a nullable date; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(DateOnly? value) =>
        value is null ? Blank : DateTime(value.Value.ToDateTime(System.TimeOnly.MinValue));

    /// <summary>Wraps a nullable time of day; <see langword="null"/> becomes <see cref="Blank"/>.</summary>
    public static implicit operator CellValue(TimeOnly? value) => value is null ? Blank : TimeSpan(value.Value.ToTimeSpan());

    /// <summary>Unwraps a number; equivalent to <see cref="GetNumber"/>.</summary>
    public static explicit operator double(CellValue value) => value.GetNumber();

    /// <summary>Unwraps text; equivalent to <see cref="GetText"/>.</summary>
    public static explicit operator string(CellValue value) => value.GetText();

    /// <summary>Unwraps a logical value; equivalent to <see cref="GetBoolean"/>.</summary>
    public static explicit operator bool(CellValue value) => value.GetBoolean();

    /// <summary>Unwraps a date-time; equivalent to <see cref="GetDateTime"/>.</summary>
    public static explicit operator DateTime(CellValue value) => value.GetDateTime();

    /// <summary>Unwraps a duration; equivalent to <see cref="GetTimeSpan"/>.</summary>
    public static explicit operator TimeSpan(CellValue value) => value.GetTimeSpan();

    /// <summary>Unwraps an error; equivalent to <see cref="GetError"/>.</summary>
    public static explicit operator ExcelError(CellValue value) => value.GetError();
}
