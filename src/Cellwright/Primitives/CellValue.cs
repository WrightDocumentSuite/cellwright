using System.Diagnostics;
using System.Globalization;

namespace Cellwright;

/// <summary>
/// The value of a single cell: a boxing-free tagged union of blank, number, text, boolean,
/// date-time, duration and error. The struct is 24 bytes; temporal values keep full tick
/// precision (they are converted to Excel serial numbers only at the file boundary).
/// </summary>
[DebuggerDisplay("{Kind}: {ToString(),nq}")]
public readonly partial struct CellValue : IEquatable<CellValue>
{
    // Number payload; boolean 0/1; error code; DateTime/TimeSpan ticks reinterpreted via Int64BitsToDouble.
    private readonly double _number;
    private readonly string? _text;

    private CellValue(CellValueKind kind, double number, string? text)
    {
        Kind = kind;
        _number = number;
        _text = text;
    }

    /// <summary>Discriminator of the stored value.</summary>
    public CellValueKind Kind { get; }

    /// <summary>A blank value — the default state of a cell.</summary>
    public static CellValue Blank => default;

    /// <summary>
    /// Creates a number value. <see cref="double.NaN"/> and infinities cannot exist in a worksheet,
    /// so they are coerced to the <c>#NUM!</c> error, mirroring what Excel formulas produce.
    /// </summary>
    public static CellValue Number(double value) =>
        double.IsFinite(value) ? new(CellValueKind.Number, value, null) : Error(ExcelError.Number);

    /// <summary>Creates a text value.</summary>
    /// <exception cref="ArgumentException">The text exceeds <see cref="SheetLimits.MaxCellTextLength"/>.</exception>
    public static CellValue Text(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (value.Length > SheetLimits.MaxCellTextLength)
            throw new ArgumentException($"Cell text cannot exceed {SheetLimits.MaxCellTextLength} characters.", nameof(value));
        return new(CellValueKind.Text, 0, value);
    }

    /// <summary>Creates a logical value.</summary>
    public static CellValue Boolean(bool value) => new(CellValueKind.Boolean, value ? 1 : 0, null);

    /// <summary>Creates a date-time value with full tick precision.</summary>
    public static CellValue DateTime(DateTime value) =>
        new(CellValueKind.DateTime, BitConverter.Int64BitsToDouble(value.Ticks), null);

    /// <summary>Creates a duration value with full tick precision.</summary>
    public static CellValue TimeSpan(TimeSpan value) =>
        new(CellValueKind.TimeSpan, BitConverter.Int64BitsToDouble(value.Ticks), null);

    /// <summary>Creates an error value.</summary>
    public static CellValue Error(ExcelError error) => new(CellValueKind.Error, (byte)error, null);

    /// <summary>Whether the value is blank.</summary>
    public bool IsBlank => Kind == CellValueKind.Blank;

    /// <summary>Whether the value is a number.</summary>
    public bool IsNumber => Kind == CellValueKind.Number;

    /// <summary>Whether the value is text.</summary>
    public bool IsText => Kind == CellValueKind.Text;

    /// <summary>Whether the value is a boolean.</summary>
    public bool IsBoolean => Kind == CellValueKind.Boolean;

    /// <summary>Whether the value is a date-time.</summary>
    public bool IsDateTime => Kind == CellValueKind.DateTime;

    /// <summary>Whether the value is a duration.</summary>
    public bool IsTimeSpan => Kind == CellValueKind.TimeSpan;

    /// <summary>Whether the value is an error.</summary>
    public bool IsError => Kind == CellValueKind.Error;

    /// <summary>Returns the number payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.Number"/>.</exception>
    public double GetNumber() => Kind == CellValueKind.Number ? _number : throw WrongKind(CellValueKind.Number);

    /// <summary>Returns the text payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.Text"/>.</exception>
    public string GetText() => Kind == CellValueKind.Text ? _text! : throw WrongKind(CellValueKind.Text);

    /// <summary>Returns the boolean payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.Boolean"/>.</exception>
    public bool GetBoolean() => Kind == CellValueKind.Boolean ? _number != 0 : throw WrongKind(CellValueKind.Boolean);

    /// <summary>Returns the date-time payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.DateTime"/>.</exception>
    public DateTime GetDateTime() => Kind == CellValueKind.DateTime
        ? new DateTime(BitConverter.DoubleToInt64Bits(_number), DateTimeKind.Unspecified)
        : throw WrongKind(CellValueKind.DateTime);

    /// <summary>Returns the duration payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.TimeSpan"/>.</exception>
    public TimeSpan GetTimeSpan() => Kind == CellValueKind.TimeSpan
        ? new TimeSpan(BitConverter.DoubleToInt64Bits(_number))
        : throw WrongKind(CellValueKind.TimeSpan);

    /// <summary>Returns the error payload.</summary>
    /// <exception cref="InvalidCastException">The value is not a <see cref="CellValueKind.Error"/>.</exception>
    public ExcelError GetError() => Kind == CellValueKind.Error ? (ExcelError)(byte)_number : throw WrongKind(CellValueKind.Error);

    /// <summary>Attempts to read the value as a number.</summary>
    public bool TryGetNumber(out double value)
    {
        value = Kind == CellValueKind.Number ? _number : 0;
        return Kind == CellValueKind.Number;
    }

    /// <summary>Attempts to read the value as text.</summary>
    public bool TryGetText(out string value)
    {
        value = Kind == CellValueKind.Text ? _text! : string.Empty;
        return Kind == CellValueKind.Text;
    }

    /// <summary>Attempts to read the value as a date-time.</summary>
    public bool TryGetDateTime(out DateTime value)
    {
        bool match = Kind == CellValueKind.DateTime;
        value = match ? GetDateTime() : default;
        return match;
    }

    /// <summary>
    /// Converts number-like payloads to the serial representation used inside xlsx files:
    /// numbers pass through, booleans become 0/1, date-times and durations become serial days.
    /// </summary>
    /// <exception cref="InvalidCastException">The value is blank, text or an error.</exception>
    public double ToSerialNumber(ExcelDateSystem system = ExcelDateSystem.Date1900) => Kind switch
    {
        CellValueKind.Number or CellValueKind.Boolean => _number,
        CellValueKind.DateTime => ExcelSerialDate.FromDateTime(GetDateTime(), system),
        CellValueKind.TimeSpan => ExcelSerialDate.FromTimeSpan(GetTimeSpan()),
        _ => throw new InvalidCastException($"A {Kind} cell value has no serial number representation."),
    };

    /// <inheritdoc/>
    public bool Equals(CellValue other) =>
        Kind == other.Kind &&
        _number.Equals(other._number) &&
        string.Equals(_text, other._text, StringComparison.Ordinal);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is CellValue other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Kind, _number, _text);

    /// <summary>Equality operator.</summary>
    public static bool operator ==(CellValue left, CellValue right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(CellValue left, CellValue right) => !left.Equals(right);

    /// <summary>Culture-invariant textual representation, mainly for logging and debugging.</summary>
    public override string ToString() => Kind switch
    {
        CellValueKind.Blank => string.Empty,
        CellValueKind.Number => _number.ToString(CultureInfo.InvariantCulture),
        CellValueKind.Text => _text!,
        CellValueKind.Boolean => _number != 0 ? "TRUE" : "FALSE",
        CellValueKind.DateTime => GetDateTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        CellValueKind.TimeSpan => GetTimeSpan().ToString(),
        CellValueKind.Error => GetError().ToText(),
        _ => string.Empty,
    };

    private InvalidCastException WrongKind(CellValueKind requested) =>
        new($"The cell value is a {Kind}, not a {requested}.");
}
