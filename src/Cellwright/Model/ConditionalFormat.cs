namespace Cellwright;

/// <summary>What a conditional formatting rule tests, mirroring the <c>cfRule</c> types.</summary>
public enum ConditionalFormatType
{
    /// <summary>Compares the cell against one or two values with <see cref="ConditionalFormatRule.Operator"/>.</summary>
    CellIs,

    /// <summary>Applies where a formula written for the top-left cell of the range is true.</summary>
    Expression,

    /// <summary>Grades the range between two or three colors.</summary>
    ColorScale,

    /// <summary>Draws a bar in each cell, proportional to its value.</summary>
    DataBar,

    /// <summary>Puts an icon in each cell according to which band its value falls in.</summary>
    IconSet,

    /// <summary>The highest or lowest values of the range.</summary>
    Top10,

    /// <summary>Values above or below the average of the range.</summary>
    AboveAverage,

    /// <summary>Text containing, or not containing, <see cref="ConditionalFormatRule.Text"/>.</summary>
    ContainsText,

    /// <inheritdoc cref="ContainsText"/>
    NotContainsText,

    /// <summary>Text starting with <see cref="ConditionalFormatRule.Text"/>.</summary>
    BeginsWith,

    /// <summary>Text ending with <see cref="ConditionalFormatRule.Text"/>.</summary>
    EndsWith,

    /// <summary>Empty cells.</summary>
    ContainsBlanks,

    /// <summary>Cells that are not empty.</summary>
    NotContainsBlanks,

    /// <summary>Cells holding an error value.</summary>
    ContainsErrors,

    /// <summary>Cells not holding an error value.</summary>
    NotContainsErrors,

    /// <summary>Values that occur more than once in the range.</summary>
    DuplicateValues,

    /// <summary>Values that occur exactly once in the range.</summary>
    UniqueValues,

    /// <summary>Dates in the period named by <see cref="ConditionalFormatRule.TimePeriod"/>.</summary>
    TimePeriod,
}

/// <summary>How a <see cref="ConditionalFormatType.CellIs"/> rule compares the cell.</summary>
public enum ConditionalOperator
{
    /// <summary>No comparison; used by the rule types that need none.</summary>
    None,

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

/// <summary>How one stop of a color scale, data bar or icon set is placed.</summary>
public enum ConditionalValueType
{
    /// <summary>The lowest value of the range.</summary>
    Minimum,

    /// <summary>The highest value of the range.</summary>
    Maximum,

    /// <summary>An absolute number.</summary>
    Number,

    /// <summary>A percentage of the span between lowest and highest.</summary>
    Percent,

    /// <summary>A percentile of the values.</summary>
    Percentile,

    /// <summary>Whatever a formula evaluates to.</summary>
    Formula,
}

/// <summary>One stop of a graded rule: where it sits and, when it needs one, its value.</summary>
/// <param name="Type">How the stop is placed.</param>
/// <param name="Value">Number or formula text, unused by <see cref="ConditionalValueType.Minimum"/> and <see cref="ConditionalValueType.Maximum"/>.</param>
public readonly record struct ConditionalValue(ConditionalValueType Type, string? Value = null)
{
    /// <summary>The lowest value of the range.</summary>
    public static ConditionalValue Minimum => new(ConditionalValueType.Minimum);

    /// <summary>The highest value of the range.</summary>
    public static ConditionalValue Maximum => new(ConditionalValueType.Maximum);

    /// <summary>An absolute number.</summary>
    public static ConditionalValue Number(double value) => new(ConditionalValueType.Number, Text(value));

    /// <summary>A percentage of the span between lowest and highest.</summary>
    public static ConditionalValue Percent(double value) => new(ConditionalValueType.Percent, Text(value));

    /// <summary>A percentile of the values.</summary>
    public static ConditionalValue Percentile(double value) => new(ConditionalValueType.Percentile, Text(value));

    /// <summary>Whatever a formula evaluates to, written without the leading <c>=</c>.</summary>
    public static ConditionalValue Formula(string formula) => new(ConditionalValueType.Formula, formula);

    private static string Text(double value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>A two- or three-color gradient over the values of a range.</summary>
public sealed record ColorScale
{
    /// <summary>Stops of the scale, low to high. Two or three of them.</summary>
    public required IReadOnlyList<ConditionalValue> Values { get; init; }

    /// <summary>Colors of the stops, one per entry of <see cref="Values"/>.</summary>
    public required IReadOnlyList<ExcelColor> Colors { get; init; }

    /// <summary>The familiar red-yellow-green scale from lowest to highest.</summary>
    public static ColorScale RedYellowGreen { get; } = new()
    {
        Values = [ConditionalValue.Minimum, ConditionalValue.Percentile(50), ConditionalValue.Maximum],
        Colors =
        [
            ExcelColor.FromRgb(0xF8, 0x69, 0x6B),
            ExcelColor.FromRgb(0xFF, 0xEB, 0x84),
            ExcelColor.FromRgb(0x63, 0xBE, 0x7B),
        ],
    };
}

/// <summary>An in-cell bar whose length tracks the value.</summary>
public sealed record DataBar
{
    /// <summary>Where the bar is empty.</summary>
    public ConditionalValue Minimum { get; init; } = ConditionalValue.Minimum;

    /// <summary>Where the bar fills the cell.</summary>
    public ConditionalValue Maximum { get; init; } = ConditionalValue.Maximum;

    /// <summary>Bar color.</summary>
    public ExcelColor Color { get; init; } = ExcelColor.FromRgb(0x63, 0x8E, 0xC6);

    /// <summary>Whether the number stays visible next to the bar.</summary>
    public bool ShowValue { get; init; } = true;
}

/// <summary>Icons chosen by which band a value falls in.</summary>
public sealed record IconSet
{
    /// <summary>Name of the built-in set, e.g. <c>3TrafficLights1</c>, <c>4Arrows</c>, <c>5Rating</c>.</summary>
    public string Name { get; init; } = "3TrafficLights1";

    /// <summary>
    /// Lower bounds of the bands, low to high. The first is always the minimum, so a set of
    /// three icons needs three entries.
    /// </summary>
    public required IReadOnlyList<ConditionalValue> Values { get; init; }

    /// <summary>Whether the number stays visible next to the icon.</summary>
    public bool ShowValue { get; init; } = true;

    /// <summary>Whether the icons run in the opposite order.</summary>
    public bool Reverse { get; init; }

    /// <summary>Red, yellow and green lights at the 33rd and 67th percent marks.</summary>
    public static IconSet TrafficLights { get; } = new()
    {
        Values = [ConditionalValue.Percent(0), ConditionalValue.Percent(33), ConditionalValue.Percent(67)],
    };
}
