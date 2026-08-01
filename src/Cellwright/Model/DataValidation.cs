using System.Collections;

namespace Cellwright;

/// <summary>What kind of input a <see cref="DataValidation"/> accepts.</summary>
public enum DataValidationType
{
    /// <summary>No restriction; useful for attaching an input message only.</summary>
    None,

    /// <summary>Whole numbers.</summary>
    Whole,

    /// <summary>Any number.</summary>
    Decimal,

    /// <summary>One of the values a formula or literal list yields.</summary>
    List,

    /// <summary>Dates.</summary>
    Date,

    /// <summary>Times.</summary>
    Time,

    /// <summary>Text of a constrained length.</summary>
    TextLength,

    /// <summary>Anything a formula accepts.</summary>
    Custom,
}

/// <summary>How the two formulas of a <see cref="DataValidation"/> bound the value.</summary>
public enum DataValidationOperator
{
    /// <summary>Between the two formulas, inclusive.</summary>
    Between,

    /// <summary>Outside the two formulas.</summary>
    NotBetween,

    /// <summary>Equal to the first formula.</summary>
    Equal,

    /// <summary>Different from the first formula.</summary>
    NotEqual,

    /// <summary>Greater than the first formula.</summary>
    GreaterThan,

    /// <summary>Less than the first formula.</summary>
    LessThan,

    /// <summary>Greater than or equal to the first formula.</summary>
    GreaterThanOrEqual,

    /// <summary>Less than or equal to the first formula.</summary>
    LessThanOrEqual,
}

/// <summary>How Excel reacts to input the validation rejects.</summary>
public enum DataValidationErrorStyle
{
    /// <summary>Refuse the input.</summary>
    Stop,

    /// <summary>Warn but allow the input.</summary>
    Warning,

    /// <summary>Inform only.</summary>
    Information,
}

/// <summary>
/// An input rule attached to one or more ranges of a worksheet: the dropdown behind a list
/// column, a numeric bound, or a custom formula, together with the messages Excel shows.
/// </summary>
public sealed class DataValidation
{
    /// <summary>Ranges the rule applies to. At least one is required.</summary>
    public required IReadOnlyList<RangeRef> Ranges { get; set; }

    /// <summary>Kind of input accepted.</summary>
    public DataValidationType Type { get; init; }

    /// <summary>How <see cref="Formula1"/> and <see cref="Formula2"/> bound the value.</summary>
    public DataValidationOperator Operator { get; init; }

    /// <summary>First bound, list source or custom condition, without the leading <c>=</c>.</summary>
    public string? Formula1 { get; set; }

    /// <summary>Second bound for the between operators.</summary>
    public string? Formula2 { get; set; }

    /// <summary>Whether an empty cell passes.</summary>
    public bool AllowBlank { get; init; } = true;

    /// <summary>Whether a list validation shows its dropdown arrow.</summary>
    public bool ShowDropDown { get; init; } = true;

    /// <summary>Whether the input message is shown while the cell is selected.</summary>
    public bool ShowInputMessage { get; init; } = true;

    /// <summary>Whether the error alert is shown on rejected input.</summary>
    public bool ShowErrorMessage { get; init; } = true;

    /// <summary>Title of the message shown while the cell is selected.</summary>
    public string? PromptTitle { get; init; }

    /// <summary>Body of the message shown while the cell is selected.</summary>
    public string? Prompt { get; init; }

    /// <summary>Title of the alert shown on rejected input.</summary>
    public string? ErrorTitle { get; init; }

    /// <summary>Body of the alert shown on rejected input.</summary>
    public string? Error { get; init; }

    /// <summary>How Excel reacts to rejected input.</summary>
    public DataValidationErrorStyle ErrorStyle { get; init; }

    /// <summary>
    /// A dropdown over a range, e.g. <c>DataValidation.List("E2:E100", "Lists!$A$1:$A$5")</c>.
    /// Literal choices go in quotes: <c>"\"Yes,No\""</c>.
    /// </summary>
    public static DataValidation List(string ranges, string source) => new()
    {
        Ranges = ParseRanges(ranges),
        Type = DataValidationType.List,
        Formula1 = source,
    };

    /// <summary>A whole-number rule bounded by two values.</summary>
    public static DataValidation WholeBetween(string ranges, long minimum, long maximum) => new()
    {
        Ranges = ParseRanges(ranges),
        Type = DataValidationType.Whole,
        Operator = DataValidationOperator.Between,
        Formula1 = minimum.ToString(System.Globalization.CultureInfo.InvariantCulture),
        Formula2 = maximum.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    /// <summary>A rule that accepts whatever a formula evaluates as true.</summary>
    public static DataValidation Custom(string ranges, string formula) => new()
    {
        Ranges = ParseRanges(ranges),
        Type = DataValidationType.Custom,
        Formula1 = formula,
    };

    /// <summary>
    /// A separate rule with the same settings, for a sheet that was duplicated. The ranges are
    /// a list of their own: a structural edit rewrites them in place.
    /// </summary>
    internal DataValidation Clone()
    {
        var copy = (DataValidation)MemberwiseClone();
        copy.Ranges = [.. Ranges];
        return copy;
    }

    /// <summary>Parses a space-separated <c>sqref</c> list such as <c>A1:A10 C1:C10</c>.</summary>
    public static IReadOnlyList<RangeRef> ParseRanges(string sqref)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sqref);
        List<RangeRef> ranges = [];
        foreach (Range segment in sqref.AsSpan().Split(' '))
        {
            ReadOnlySpan<char> text = sqref.AsSpan()[segment];
            if (!text.IsEmpty && RangeRef.TryParse(text, null, out RangeRef range))
                ranges.Add(range);
        }

        if (ranges.Count == 0)
            throw new FormatException($"'{sqref}' contains no valid range reference.");
        return ranges;
    }

    /// <inheritdoc/>
    public override string ToString() => $"{Type} over {string.Join(' ', Ranges)}";
}

/// <summary>The data validation rules of one worksheet.</summary>
public sealed class DataValidationCollection : IReadOnlyList<DataValidation>
{
    private readonly List<DataValidation> _validations = [];

    /// <inheritdoc/>
    public int Count => _validations.Count;

    /// <inheritdoc/>
    public DataValidation this[int index] => _validations[index];

    /// <summary>Whether the model, rather than the loaded file, owns the sheet's validations.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>Adds a rule.</summary>
    public DataValidation Add(DataValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        if (validation.Ranges.Count == 0)
            throw new ArgumentException("A data validation must cover at least one range.", nameof(validation));

        _validations.Add(validation);
        IsModified = true;
        return validation;
    }

    /// <summary>Removes a rule. Returns whether anything was removed.</summary>
    public bool Remove(DataValidation validation)
    {
        bool removed = _validations.Remove(validation);
        IsModified |= removed;
        return removed;
    }

    /// <summary>Removes every rule.</summary>
    public void Clear()
    {
        if (_validations.Count == 0)
            return;
        _validations.Clear();
        IsModified = true;
    }

    /// <summary>Adds a rule read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(DataValidation validation) => _validations.Add(validation);

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit moved a rule.</summary>
    internal void MarkModified() => IsModified = true;

    /// <inheritdoc/>
    public IEnumerator<DataValidation> GetEnumerator() => _validations.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
