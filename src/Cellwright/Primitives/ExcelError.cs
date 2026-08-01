using System.Collections.Frozen;

namespace Cellwright;

/// <summary>
/// Excel cell error values, including the dynamic-array era errors.
/// </summary>
public enum ExcelError : byte
{
    /// <summary><c>#NULL!</c> — ranges do not intersect.</summary>
    Null = 0,

    /// <summary><c>#DIV/0!</c> — division by zero.</summary>
    DivisionByZero = 1,

    /// <summary><c>#VALUE!</c> — wrong operand type.</summary>
    Value = 2,

    /// <summary><c>#REF!</c> — invalid cell reference.</summary>
    Reference = 3,

    /// <summary><c>#NAME?</c> — unrecognized defined name or function.</summary>
    Name = 4,

    /// <summary><c>#NUM!</c> — invalid numeric value.</summary>
    Number = 5,

    /// <summary><c>#N/A</c> — value not available.</summary>
    NotAvailable = 6,

    /// <summary><c>#GETTING_DATA</c> — asynchronous data is still loading.</summary>
    GettingData = 7,

    /// <summary><c>#SPILL!</c> — dynamic array result is blocked.</summary>
    Spill = 8,

    /// <summary><c>#CALC!</c> — calculation engine error.</summary>
    Calc = 9,
}

/// <summary>
/// Conversions between <see cref="ExcelError"/> and the literal error strings used in xlsx files.
/// </summary>
public static class ExcelErrorText
{
    private static readonly FrozenDictionary<string, ExcelError> Lookup = new Dictionary<string, ExcelError>(StringComparer.Ordinal)
    {
        ["#NULL!"] = ExcelError.Null,
        ["#DIV/0!"] = ExcelError.DivisionByZero,
        ["#VALUE!"] = ExcelError.Value,
        ["#REF!"] = ExcelError.Reference,
        ["#NAME?"] = ExcelError.Name,
        ["#NUM!"] = ExcelError.Number,
        ["#N/A"] = ExcelError.NotAvailable,
        ["#GETTING_DATA"] = ExcelError.GettingData,
        ["#SPILL!"] = ExcelError.Spill,
        ["#CALC!"] = ExcelError.Calc,
    }.ToFrozenDictionary();

    private static readonly FrozenDictionary<string, ExcelError>.AlternateLookup<ReadOnlySpan<char>> SpanLookup =
        Lookup.GetAlternateLookup<ReadOnlySpan<char>>();

    /// <summary>Returns the literal Excel representation, e.g. <c>#DIV/0!</c>.</summary>
    public static string ToText(this ExcelError error) => error switch
    {
        ExcelError.Null => "#NULL!",
        ExcelError.DivisionByZero => "#DIV/0!",
        ExcelError.Value => "#VALUE!",
        ExcelError.Reference => "#REF!",
        ExcelError.Name => "#NAME?",
        ExcelError.Number => "#NUM!",
        ExcelError.NotAvailable => "#N/A",
        ExcelError.GettingData => "#GETTING_DATA",
        ExcelError.Spill => "#SPILL!",
        ExcelError.Calc => "#CALC!",
        _ => throw new ArgumentOutOfRangeException(nameof(error)),
    };

    /// <summary>Tries to interpret <paramref name="text"/> as a literal Excel error string.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out ExcelError error) =>
        SpanLookup.TryGetValue(text, out error);
}
