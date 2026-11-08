using System.Collections;

namespace Cellwright;

/// <summary>
/// One conditional formatting rule: what it tests, over which ranges, and how the cells that
/// pass are painted. Rules of a sheet are tried in <see cref="Priority"/> order.
/// </summary>
/// <example>
/// <code>
/// sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B2:B99", 1000, DifferentialStyle.Good));
/// sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("C2:C99", ColorScale.RedYellowGreen));
/// </code>
/// </example>
public sealed class ConditionalFormatRule
{
    /// <summary>Ranges the rule covers. At least one is required.</summary>
    public required IReadOnlyList<RangeRef> Ranges
    {
        get => _ranges;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            _ranges = [.. value];
            Owner?.MarkModified();
        }
    }

    private IReadOnlyList<RangeRef> _ranges = [];

    /// <summary>What the rule tests.</summary>
    public ConditionalFormatType Type { get; init; }

    /// <summary>Comparison of a <see cref="ConditionalFormatType.CellIs"/> rule.</summary>
    public ConditionalOperator Operator { get; init; }

    /// <summary>
    /// Formulas the rule needs, without the leading <c>=</c>: the bounds of a comparison or the
    /// single condition of an <see cref="ConditionalFormatType.Expression"/> rule; graded rules
    /// can use the first formula as their activity condition. Relative references are written as
    /// they read in the top-left cell of the first range.
    /// </summary>
    public IReadOnlyList<string> Formulas
    {
        get => _formulas;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            _formulas = [.. value];
            Owner?.MarkModified();
        }
    }

    private IReadOnlyList<string> _formulas = [];

    /// <summary>Needle of the text rules.</summary>
    public string? Text { get; init; }

    /// <summary>How many values a <see cref="ConditionalFormatType.Top10"/> rule takes.</summary>
    public int Rank { get; init; } = 10;

    /// <summary>Whether a <see cref="ConditionalFormatType.Top10"/> rule takes the lowest values.</summary>
    public bool Bottom { get; init; }

    /// <summary>Whether <see cref="Rank"/> is a percentage rather than a count.</summary>
    public bool Percent { get; init; }

    /// <summary>Whether an <see cref="ConditionalFormatType.AboveAverage"/> rule looks above rather than below.</summary>
    public bool Above { get; init; } = true;

    /// <summary>Whether values exactly on the average count as passing.</summary>
    public bool EqualAverage { get; init; }

    /// <summary>How many standard deviations from the average the boundary sits, or 0 for the average itself.</summary>
    public int StandardDeviation { get; init; }

    /// <summary>Period of a <see cref="ConditionalFormatType.TimePeriod"/> rule, e.g. <c>lastWeek</c>, <c>today</c>.</summary>
    public string? TimePeriod { get; init; }

    /// <summary>Order in which the sheet's rules are tried; the lowest number wins first.</summary>
    public int Priority
    {
        get => _priority;
        set
        {
            if (_priority == value)
                return;
            _priority = value;
            Owner?.MarkModified();
        }
    }

    private int _priority;

    /// <summary>Whether a match stops the rules after this one from being tried.</summary>
    public bool StopIfTrue { get; init; }

    /// <summary>
    /// How matching cells are painted. Graded rules paint themselves and leave this null. For a
    /// rule read from a file this shows the parsed view of its differential format; assigning a
    /// different one hands the sheet's rules to the model, like any other edit.
    /// </summary>
    public DifferentialStyle? Style
    {
        get => _style;
        set
        {
            if (_style == value)
                return;
            _style = value;
            Owner?.MarkModified();
        }
    }

    private DifferentialStyle? _style;

    /// <summary>The collection the rule sits in, so edits to the rule can mark it as owned.</summary>
    internal ConditionalFormatCollection? Owner { get; set; }

    /// <summary>Shows the rule its file's differential format without claiming an edit happened.</summary>
    internal void AttachSourceStyle(DifferentialStyle style) => _style = style;

    /// <summary>Gradient of a <see cref="ConditionalFormatType.ColorScale"/> rule.</summary>
    public ColorScale? ColorScale { get; init; }

    /// <summary>Bar of a <see cref="ConditionalFormatType.DataBar"/> rule.</summary>
    public DataBar? DataBar { get; init; }

    /// <summary>Icons of a <see cref="ConditionalFormatType.IconSet"/> rule.</summary>
    public IconSet? IconSet { get; init; }

    /// <summary>
    /// Differential format the loaded file pointed at. Used when <see cref="Style"/> is null so
    /// that a rule read from a file keeps the exact formatting it had, which the model does not
    /// otherwise parse.
    /// </summary>
    internal int? SourceStyleId { get; init; }

    /// <summary>
    /// Verbatim <c>cfRule</c> element of a rule the model cannot represent — a rule carrying x14
    /// extensions, say. It is replayed unchanged, so editing the rules around it loses nothing.
    /// </summary>
    internal string? RawXml { get; init; }

    /// <summary>Whether the rule is only carried through, with nothing of it modelled.</summary>
    public bool IsOpaque => RawXml is not null;

    /// <summary>
    /// A separate rule with the same settings, for a sheet that was duplicated. The ranges and
    /// formulas become lists of their own: a structural edit rewrites both in place.
    /// </summary>
    internal ConditionalFormatRule Clone()
    {
        var copy = (ConditionalFormatRule)MemberwiseClone();
        copy.Owner = null;
        copy._ranges = [.. _ranges];
        copy._formulas = [.. _formulas];
        return copy;
    }

    /// <summary>Paints cells greater than a number.</summary>
    public static ConditionalFormatRule CellGreaterThan(string ranges, double value, DifferentialStyle style) =>
        Comparison(ranges, ConditionalOperator.GreaterThan, style, Number(value));

    /// <summary>Paints cells less than a number.</summary>
    public static ConditionalFormatRule CellLessThan(string ranges, double value, DifferentialStyle style) =>
        Comparison(ranges, ConditionalOperator.LessThan, style, Number(value));

    /// <summary>Paints cells between two numbers, inclusive.</summary>
    public static ConditionalFormatRule CellBetween(string ranges, double low, double high, DifferentialStyle style) =>
        Comparison(ranges, ConditionalOperator.Between, style, Number(low), Number(high));

    /// <summary>Paints cells equal to a number.</summary>
    public static ConditionalFormatRule CellEqualTo(string ranges, double value, DifferentialStyle style) =>
        Comparison(ranges, ConditionalOperator.Equal, style, Number(value));

    /// <summary>
    /// Paints cells where a formula is true. Write it as it reads in the top-left cell of the
    /// range: <c>Expression("A2:D99", "$D2&gt;100", …)</c> highlights whole rows.
    /// </summary>
    public static ConditionalFormatRule Expression(string ranges, string formula, DifferentialStyle style) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.Expression,
        Formulas = [formula],
        Style = style,
    };

    /// <summary>Paints cells containing a piece of text.</summary>
    public static ConditionalFormatRule Containing(string ranges, string text, DifferentialStyle style) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.ContainsText,
        Operator = ConditionalOperator.None,
        Text = text,
        Style = style,
    };

    /// <summary>Paints values that occur more than once in the range.</summary>
    public static ConditionalFormatRule Duplicates(string ranges, DifferentialStyle style) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.DuplicateValues,
        Style = style,
    };

    /// <summary>Paints the highest values of the range.</summary>
    public static ConditionalFormatRule Top(string ranges, int count, DifferentialStyle style) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.Top10,
        Rank = count,
        Style = style,
    };

    /// <summary>Grades the range with a color scale.</summary>
    public static ConditionalFormatRule Scale(string ranges, ColorScale scale) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.ColorScale,
        ColorScale = scale,
    };

    /// <summary>Draws a bar in each cell of the range.</summary>
    public static ConditionalFormatRule Bars(string ranges, DataBar bar) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.DataBar,
        DataBar = bar,
    };

    /// <summary>Puts an icon in each cell of the range.</summary>
    public static ConditionalFormatRule Icons(string ranges, IconSet icons) => new()
    {
        Ranges = DataValidation.ParseRanges(ranges),
        Type = ConditionalFormatType.IconSet,
        IconSet = icons,
    };

    private static ConditionalFormatRule Comparison(
        string ranges, ConditionalOperator op, DifferentialStyle style, params string[] formulas) => new()
        {
            Ranges = DataValidation.ParseRanges(ranges),
            Type = ConditionalFormatType.CellIs,
            Operator = op,
            Formulas = formulas,
            Style = style,
        };

    private static string Number(double value) =>
        value.ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <inheritdoc/>
    public override string ToString() => $"{Type} over {string.Join(' ', Ranges)}";
}

/// <summary>
/// The conditional formatting rules of one worksheet, in the order they were read or added.
/// While nothing here is touched, the sheet's original markup is written back unchanged; the
/// first edit switches to generating every rule from this model.
/// </summary>
public sealed class ConditionalFormatCollection : IReadOnlyList<ConditionalFormatRule>
{
    private readonly List<ConditionalFormatRule> _rules = [];

    /// <inheritdoc/>
    public int Count => _rules.Count;

    /// <inheritdoc/>
    public ConditionalFormatRule this[int index] => _rules[index];

    /// <summary>Whether the model, rather than the loaded file, owns the sheet's rules.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>Adds a rule, giving it the next priority when it has none.</summary>
    public ConditionalFormatRule Add(ConditionalFormatRule rule)
    {
        ArgumentNullException.ThrowIfNull(rule);
        ThrowIfOwned(rule);
        if (rule.Ranges.Count == 0)
            throw new ArgumentException("A conditional format must cover at least one range.", nameof(rule));
        if (rule.DataBar is { HasValidLengths: false })
        {
            throw new ArgumentException(
                "Data-bar lengths must satisfy 0 <= MinimumLength <= MaximumLength <= 100.",
                nameof(rule));
        }

        if (rule.Priority <= 0 || PriorityTaken(rule.Priority))
            rule.Priority = NextPriority();
        rule.Owner = this;
        _rules.Add(rule);
        IsModified = true;
        return rule;
    }

    /// <summary>Removes a rule. Returns whether anything was removed.</summary>
    public bool Remove(ConditionalFormatRule rule)
    {
        bool removed = _rules.Remove(rule);
        if (removed)
        {
            rule.Owner = null;
            IsModified = true;
        }

        return removed;
    }

    /// <summary>Removes every rule.</summary>
    public void Clear()
    {
        if (_rules.Count == 0)
            return;

        foreach (ConditionalFormatRule rule in _rules)
            rule.Owner = null;
        _rules.Clear();
        IsModified = true;
    }

    /// <summary>Adds a rule read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(ConditionalFormatRule rule)
    {
        ArgumentNullException.ThrowIfNull(rule);
        ThrowIfOwned(rule);
        rule.Owner = this;
        _rules.Add(rule);
    }

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit.</summary>
    internal void MarkModified() => IsModified = true;

    private static void ThrowIfOwned(ConditionalFormatRule rule)
    {
        if (rule.Owner is not null)
        {
            throw new ArgumentException(
                "A conditional-format rule already belongs to a worksheet. Remove it first or add a separate rule instance.",
                nameof(rule));
        }
    }

    private bool PriorityTaken(int priority)
    {
        foreach (ConditionalFormatRule existing in _rules)
        {
            if (existing.Priority == priority)
                return true;
        }

        return false;
    }

    private int NextPriority()
    {
        int highest = 0;
        foreach (ConditionalFormatRule rule in _rules)
            highest = Math.Max(highest, rule.Priority);
        if (highest < int.MaxValue)
            return highest + 1;

        // A malformed or adversarial input can already use Int32.MaxValue. Do not wrap the
        // next rule to a negative priority: find the first positive gap instead.
        var used = new HashSet<int>(_rules.Count);
        foreach (ConditionalFormatRule rule in _rules)
        {
            if (rule.Priority > 0)
                used.Add(rule.Priority);
        }

        for (int candidate = 1; candidate < int.MaxValue; candidate++)
        {
            if (!used.Contains(candidate))
                return candidate;
        }

        throw new InvalidOperationException("No positive conditional-format priority is available.");
    }

    /// <inheritdoc/>
    public IEnumerator<ConditionalFormatRule> GetEnumerator() => _rules.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
