using System.Globalization;

namespace Cellwright;

/// <summary>The bar a data-bar rule draws in a cell.</summary>
/// <param name="Fraction">How much of the cell it fills, 0 through 1.</param>
/// <param name="Color">The colour of the bar.</param>
/// <param name="ShowValue">Whether the number stays visible beside it.</param>
internal readonly record struct ConditionalBar(double Fraction, ExcelColor Color, bool ShowValue);

/// <summary>The icon an icon-set rule puts in a cell.</summary>
/// <param name="SetName">Name of the built-in set, e.g. <c>3TrafficLights1</c>.</param>
/// <param name="Index">Which icon of the set, counted from the lowest band.</param>
/// <param name="Count">How many icons the set holds.</param>
/// <param name="ShowValue">Whether the number stays visible beside it.</param>
internal readonly record struct ConditionalIcon(string SetName, int Index, int Count, bool ShowValue);

/// <summary>
/// What the sheet's conditional formatting comes to for one cell.
/// </summary>
internal readonly record struct ConditionalAppearance
{
    /// <summary>The differential formats that matched, already merged, or null when none did.</summary>
    public DifferentialStyle? Style { get; init; }

    /// <summary>The colour a scale graded this cell to, or <see cref="ExcelColor.Auto"/>.</summary>
    public ExcelColor ScaleColor { get; init; }

    /// <summary>The bar to draw, when a rule asks for one.</summary>
    public ConditionalBar? Bar { get; init; }

    /// <summary>The icon to draw, when a rule asks for one.</summary>
    public ConditionalIcon? Icon { get; init; }

    /// <summary>Whether no rule painted this cell at all.</summary>
    public bool IsEmpty => Style is null && ScaleColor.IsAuto && Bar is null && Icon is null;
}

/// <summary>
/// Works out what a sheet's conditional formatting rules do to its cells.
/// </summary>
/// <remarks>
/// This lives in the core rather than in the printing package because none of it is about
/// drawing: it is what the rules mean. Almost every rule type can be judged from the values
/// alone — the comparisons against literals, the text tests, the blanks and errors, the
/// duplicates, the ranks, the averages and every graded rule. Only the rules that hold a real
/// formula need more, and for those the caller supplies an evaluator. Without one they are
/// reported rather than guessed at, because a rule quietly treated as false is a cell that looks
/// deliberately unpainted.
/// </remarks>
internal sealed class ConditionalFormatEvaluator
{
    private readonly Worksheet _sheet;
    private readonly Func<string, CellRef, CellValue>? _evaluate;
    private readonly Func<string, CellRef, CellRef, CellValue>? _evaluateRelative;
    private readonly DateOnly _today;
    private readonly List<ConditionalFormatRule> _rules = [];
    private readonly Dictionary<ConditionalFormatRule, ConditionalRangeStatistics> _statistics = [];
    private readonly Dictionary<(CellStyle Stated, DifferentialStyle Format), CellStyle> _overlaid = [];
    private readonly Dictionary<(ConditionalFormatRule Rule, int Index), CellValue?> _operands = [];
    private readonly Dictionary<(ConditionalFormatRule Rule, string Formula), double?> _stopParameters = [];
    private readonly HashSet<string> _unsupported = [];

    /// <summary>Prepares the rules of a sheet for evaluation.</summary>
    /// <param name="sheet">The sheet whose rules to apply.</param>
    /// <param name="evaluate">
    /// How to evaluate a rule's formula in a cell, or <see langword="null"/> when the caller has
    /// no formula engine. Reference <c>Cellwright.Formulas</c> and pass
    /// <c>(formula, cell) =&gt; sheet.Evaluate(formula, cell.ToString())</c>.
    /// Relative references have already been copied from the first range's top-left cell to the
    /// target cell before this delegate is called.
    /// </param>
    /// <param name="today">What counts as today for the rules about time periods.</param>
    /// <param name="evaluateRelative">
    /// Cache-stable counterpart of <paramref name="evaluate"/>: it receives the formula as stored,
    /// its anchor, and its target. A formula engine can compile the stored text once and apply the
    /// copy offset during evaluation instead of caching one translated string per cell.
    /// </param>
    public ConditionalFormatEvaluator(
        Worksheet sheet,
        Func<string, CellRef, CellValue>? evaluate = null,
        DateOnly? today = null,
        Func<string, CellRef, CellRef, CellValue>? evaluateRelative = null)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        _sheet = sheet;
        _evaluate = evaluate;
        _evaluateRelative = evaluateRelative;
        _today = today ?? DateOnly.FromDateTime(DateTime.Today);

        if (sheet.Markup?.HasUnsupportedExtensionConditionalFormats is true)
            _unsupported.Add("conditional-format-opaque");

        foreach (ConditionalFormatRule rule in sheet.ConditionalFormats)
        {
            // A rule the model could not represent is carried through the file verbatim; there
            // is nothing here to apply it from.
            if (rule.IsOpaque)
                _unsupported.Add("conditional-format-opaque");
            else
                _rules.Add(rule);
        }

        // Supported x14 rules are a read-only semantic projection over the preserved extLst.
        // They never enter the public collection, otherwise a later base-rule edit would write
        // them in the base namespace and then replay the original extension a second time.
        if (sheet.Markup?.ExtensionConditionalFormats is { } extension)
        {
            foreach (ConditionalFormatRule rule in extension.Rules)
                _rules.Add(rule);
        }

        StableSortByPriority(_rules);
    }

    /// <summary>
    /// Orders rules without disturbing their collection order when a malformed or manually built
    /// model gives two of them the same priority. Binary formats require unique priorities, but a
    /// deterministic fallback is safer than letting <c>List.Sort</c> choose a winner.
    /// </summary>
    private static void StableSortByPriority(List<ConditionalFormatRule> rules)
    {
        // Enumerable.OrderBy is stable and keeps the adversarial reverse-priority case at
        // O(n log n). A quadratic insertion pass is visible on imported sheets with thousands
        // of rules and would turn PDF composition into an avoidable denial-of-service surface.
        ConditionalFormatRule[] ordered = [.. rules.OrderBy(static rule => rule.Priority)];
        rules.Clear();
        rules.AddRange(ordered);
    }

    /// <summary>Whether the sheet has any rule this can apply.</summary>
    public bool HasRules => _rules.Count > 0;

    /// <summary>
    /// The kinds of rule that were met and could not be applied, as diagnostic codes. Empty until
    /// something has actually been evaluated.
    /// </summary>
    public IReadOnlyCollection<string> Unsupported => _unsupported;

    /// <summary>What the rules come to for one cell.</summary>
    /// <param name="cell">The cell to judge.</param>
    public ConditionalAppearance Of(CellRef cell)
    {
        if (_rules.Count == 0)
            return default;

        CellValue value = _sheet.Values.Get(cell.Row, cell.Column);

        DifferentialStyle? style = null;
        ExcelColor scale = ExcelColor.Auto;
        ConditionalBar? bar = null;
        ConditionalIcon? icon = null;

        foreach (ConditionalFormatRule rule in _rules)
        {
            if (!Covers(rule, cell))
                continue;

            bool stop = false;
            switch (rule.Type)
            {
                case ConditionalFormatType.ColorScale when rule.ColorScale is { } colorScale:
                {
                    if (rule.Formulas.Count > 0 && !IsFormulaTrue(rule, cell, 0))
                        break;

                    ExcelColor graded = Grade(rule, colorScale, value);
                    if (!graded.IsAuto)
                    {
                        // Rules arrive highest-priority first. A scale owns the fill channel, but
                        // only if an earlier differential format has not already claimed it.
                        if (scale.IsAuto && (style is null || style.FillColor.IsAuto))
                            scale = graded;
                        stop = rule.StopIfTrue;
                    }
                    break;
                }

                case ConditionalFormatType.DataBar when rule.DataBar is { } dataBar:
                {
                    if (rule.Formulas.Count > 0 && !IsFormulaTrue(rule, cell, 0))
                        break;

                    ConditionalBar? measured = Measure(rule, dataBar, value);
                    if (measured is not null)
                    {
                        bar ??= measured;
                        stop = rule.StopIfTrue;
                    }
                    break;
                }

                case ConditionalFormatType.IconSet when rule.IconSet is { } iconSet:
                {
                    if (rule.Formulas.Count > 0 && !IsFormulaTrue(rule, cell, 0))
                        break;

                    ConditionalIcon? chosen = Choose(rule, iconSet, value);
                    if (chosen is not null)
                    {
                        icon ??= chosen;
                        stop = rule.StopIfTrue;
                    }
                    break;
                }

                default:
                    if (!Matches(rule, cell, value))
                        continue;

                    // Lower-priority rules fill only properties the earlier rules did not name.
                    if (rule.Style is { IsEmpty: false } matched)
                        style = style is null ? matched : Merge(style, matched);
                    stop = rule.StopIfTrue;
                    break;
            }

            if (stop)
                break;
        }

        return new ConditionalAppearance
        {
            Style = style,
            ScaleColor = scale,
            Bar = bar,
            Icon = icon,
        };
    }

    /// <summary>
    /// A cell's stated style with what the rules came to laid over it.
    /// </summary>
    /// <param name="stated">The style the cell wears of its own.</param>
    /// <param name="appearance">What the rules came to for it.</param>
    /// <remarks>
    /// Kept behind a cache because the answer depends only on the pair, and the pairs are few:
    /// a sheet has a handful of styles and a handful of rules, however many cells it has. Without
    /// it a column of a hundred thousand conditioned cells allocates a hundred thousand identical
    /// style records — and does it once per drawing pass.
    /// </remarks>
    public CellStyle Overlay(CellStyle stated, in ConditionalAppearance appearance)
    {
        if (appearance.Style is not { } format)
            return stated;

        (CellStyle, DifferentialStyle) key = (stated, format);
        if (_overlaid.TryGetValue(key, out CellStyle? cached))
            return cached;

        CellStyle result = format.ApplyTo(stated);
        _overlaid[key] = result;
        return result;
    }

    private static bool Covers(ConditionalFormatRule rule, CellRef cell)
    {
        foreach (RangeRef range in rule.Ranges)
        {
            if (range.Contains(cell))
                return true;
        }

        return false;
    }

    /// <summary>Whether a predicate rule holds for a cell.</summary>
    private bool Matches(ConditionalFormatRule rule, CellRef cell, in CellValue value)
    {
        switch (rule.Type)
        {
            case ConditionalFormatType.ContainsBlanks:
                return IsBlank(value);
            case ConditionalFormatType.NotContainsBlanks:
                return !IsBlank(value);
            case ConditionalFormatType.ContainsErrors:
                return value.Kind == CellValueKind.Error;
            case ConditionalFormatType.NotContainsErrors:
                return value.Kind != CellValueKind.Error;

            case ConditionalFormatType.ContainsText:
                return Text(value).Contains(rule.Text ?? "", StringComparison.OrdinalIgnoreCase);
            case ConditionalFormatType.NotContainsText:
                return !Text(value).Contains(rule.Text ?? "", StringComparison.OrdinalIgnoreCase);
            case ConditionalFormatType.BeginsWith:
                return Text(value).StartsWith(rule.Text ?? "", StringComparison.OrdinalIgnoreCase);
            case ConditionalFormatType.EndsWith:
                return Text(value).EndsWith(rule.Text ?? "", StringComparison.OrdinalIgnoreCase);

            case ConditionalFormatType.DuplicateValues:
                return !value.IsBlank && Statistics(rule, occurrences: true).OccurrencesOf(value) > 1;
            case ConditionalFormatType.UniqueValues:
                return !value.IsBlank && Statistics(rule, occurrences: true).OccurrencesOf(value) == 1;

            case ConditionalFormatType.Top10:
                return Ranks(rule, value);
            case ConditionalFormatType.AboveAverage:
                return CompareToAverage(rule, value);
            case ConditionalFormatType.TimePeriod:
                return InPeriod(rule, value);

            case ConditionalFormatType.CellIs:
                return Compares(rule, cell, value);
            case ConditionalFormatType.Expression:
                return IsFormulaTrue(rule, cell, 0);

            default:
                _unsupported.Add($"conditional-format-{rule.Type}");
                return false;
        }
    }

    /// <summary>A cell-comparison rule against one or two operands.</summary>
    private bool Compares(ConditionalFormatRule rule, CellRef cell, in CellValue value)
    {
        if (rule.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
        {
            if (!TryOperand(rule, cell, 0, out CellValue low) || !TryOperand(rule, cell, 1, out CellValue high))
                return false;

            int belowLow = Compare(value, low);
            int aboveHigh = Compare(value, high);
            if (belowLow == int.MinValue || aboveHigh == int.MinValue)
                return false;

            bool between = belowLow >= 0 && aboveHigh <= 0;
            return rule.Operator == ConditionalOperator.Between ? between : !between;
        }

        if (!TryOperand(rule, cell, 0, out CellValue operand))
            return false;

        int order = Compare(value, operand);
        if (order == int.MinValue)
            return false;

        return rule.Operator switch
        {
            ConditionalOperator.Equal => order == 0,
            ConditionalOperator.NotEqual => order != 0,
            ConditionalOperator.GreaterThan => order > 0,
            ConditionalOperator.LessThan => order < 0,
            ConditionalOperator.GreaterThanOrEqual => order >= 0,
            ConditionalOperator.LessThanOrEqual => order <= 0,
            _ => false,
        };
    }

    /// <summary>
    /// One operand of a comparison. A literal number or quoted string is read straight off;
    /// anything else is a formula and needs the engine.
    /// </summary>
    private bool TryOperand(ConditionalFormatRule rule, CellRef cell, int index, out CellValue operand)
    {
        operand = CellValue.Blank;
        if (index >= rule.Formulas.Count)
            return false;

        // Whether an operand is a literal, and what it comes to, depends on the rule alone. Every
        // cell the rule covers would otherwise re-trim and re-parse the same string.
        if (!_operands.TryGetValue((rule, index), out CellValue? literal))
        {
            literal = ParseLiteral(rule.Formulas[index]);
            _operands[(rule, index)] = literal;
        }

        if (literal is { } parsed)
        {
            operand = parsed;
            return true;
        }

        if (!CanEvaluateFormula)
        {
            _unsupported.Add("conditional-format-formula");
            return false;
        }

        operand = EvaluateFormula(rule, cell, index);
        return true;
    }

    /// <summary>
    /// An operand as the value it stands for, or null when it is a formula rather than a literal.
    /// </summary>
    private static CellValue? ParseLiteral(string source)
    {
        ReadOnlySpan<char> text = source.AsSpan().Trim();
        if (text.Length == 0)
            return CellValue.Blank;

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            return CellValue.Number(number);

        if (text.Length >= 2 && text[0] == '"' && text[^1] == '"')
            return CellValue.Text(text[1..^1].ToString().Replace("\"\"", "\"", StringComparison.Ordinal));

        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            return CellValue.Boolean(true);
        if (text.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return CellValue.Boolean(false);

        return null;
    }

    /// <summary>Whether a rule's formula evaluates to something Excel counts as true.</summary>
    private bool IsFormulaTrue(ConditionalFormatRule rule, CellRef cell, int index)
    {
        if (index >= rule.Formulas.Count)
            return false;

        if (!CanEvaluateFormula)
        {
            _unsupported.Add("conditional-format-formula");
            return false;
        }

        CellValue result = EvaluateFormula(rule, cell, index);
        return result.Kind switch
        {
            CellValueKind.Boolean => result.GetBoolean(),
            CellValueKind.Number => result.GetNumber() != 0,
            _ => false,
        };
    }

    /// <summary>
    /// Evaluates a rule formula as copied from the first range's top-left cell to the cell being
    /// judged. The callback still receives the target as its origin so functions such as ROW()
    /// and COLUMN() see the right cell.
    /// </summary>
    private CellValue EvaluateFormula(ConditionalFormatRule rule, CellRef cell, int index)
    {
        CellRef anchor = rule.Ranges.Count == 0 ? cell : rule.Ranges[0].Start;
        if (_evaluateRelative is { } evaluateRelative)
            return evaluateRelative(rule.Formulas[index], anchor, cell);

        string translated = FormulaText.Translate(rule.Formulas[index], anchor, cell)!;
        return _evaluate!(translated, cell);
    }

    private bool CanEvaluateFormula => _evaluate is not null || _evaluateRelative is not null;

    private bool Ranks(ConditionalFormatRule rule, in CellValue value)
    {
        if (!ConditionalRangeStatistics.TryNumber(value, DateSystemOf(), out double number))
            return false;

        ConditionalRangeStatistics statistics = Statistics(rule, occurrences: false);
        if (statistics.Count == 0)
            return false;

        double threshold = statistics.RankThreshold(rule.Rank, rule.Bottom, rule.Percent);
        return rule.Bottom ? number <= threshold : number >= threshold;
    }

    private bool CompareToAverage(ConditionalFormatRule rule, in CellValue value)
    {
        if (!ConditionalRangeStatistics.TryNumber(value, DateSystemOf(), out double number))
            return false;

        ConditionalRangeStatistics statistics = Statistics(rule, occurrences: false);
        if (statistics.Count == 0)
            return false;

        double boundary = statistics.Average +
            (rule.StandardDeviation * statistics.StandardDeviation * (rule.Above ? 1 : -1));

        if (rule.EqualAverage)
            return rule.Above ? number >= boundary : number <= boundary;

        return rule.Above ? number > boundary : number < boundary;
    }

    private bool InPeriod(ConditionalFormatRule rule, in CellValue value)
    {
        if (!TryDate(value, out DateOnly date))
            return false;

        int days = date.DayNumber - _today.DayNumber;
        return rule.TimePeriod switch
        {
            "today" => days == 0,
            "yesterday" => days == -1,
            "tomorrow" => days == 1,
            "last7Days" => days <= 0 && days >= -6,
            "thisWeek" => IsSameWeek(date, 0),
            "lastWeek" => IsSameWeek(date, -1),
            "nextWeek" => IsSameWeek(date, 1),
            "thisMonth" => date.Year == _today.Year && date.Month == _today.Month,
            "lastMonth" => IsMonthOffset(date, -1),
            "nextMonth" => IsMonthOffset(date, 1),
            _ => Unknown(rule.TimePeriod),
        };

        bool Unknown(string? period)
        {
            _unsupported.Add($"conditional-format-period-{period}");
            return false;
        }
    }

    /// <summary>Whether a date falls in the week that is a given number of weeks from this one.</summary>
    /// <remarks>Weeks run Sunday to Saturday, which is where Excel puts the boundary.</remarks>
    private bool IsSameWeek(DateOnly date, int offset)
    {
        DateOnly startOfThisWeek = _today.AddDays(-(int)_today.DayOfWeek);
        DateOnly start = startOfThisWeek.AddDays(offset * 7);
        return date >= start && date < start.AddDays(7);
    }

    private bool IsMonthOffset(DateOnly date, int offset)
    {
        DateOnly moved = new DateOnly(_today.Year, _today.Month, 1).AddMonths(offset);
        return date.Year == moved.Year && date.Month == moved.Month;
    }

    private bool TryDate(in CellValue value, out DateOnly date)
    {
        if (value.Kind == CellValueKind.DateTime)
        {
            date = DateOnly.FromDateTime(value.GetDateTime());
            return true;
        }

        if (value.Kind == CellValueKind.Number)
        {
            double serial = value.GetNumber();
            if (serial >= 1 && serial < 2958466)
            {
                date = DateOnly.FromDateTime(ExcelSerialDate.ToDateTime(serial, DateSystemOf()));
                return true;
            }
        }

        date = default;
        return false;
    }

    /// <summary>The colour a scale grades a value to.</summary>
    private ExcelColor Grade(ConditionalFormatRule rule, ColorScale scale, in CellValue value)
    {
        if (scale.Values.Count < 2 || scale.Colors.Count < scale.Values.Count)
            return ExcelColor.Auto;

        if (!ConditionalRangeStatistics.TryNumber(value, DateSystemOf(), out double number))
            return ExcelColor.Auto;

        ConditionalRangeStatistics statistics = Statistics(rule, occurrences: false);
        int stops = scale.Values.Count;

        Span<double> positions = stops <= 4 ? stackalloc double[stops] : new double[stops];
        for (int i = 0; i < stops; i++)
            positions[i] = Position(scale.Values[i], statistics, rule);

        if (number <= positions[0])
            return scale.Colors[0];
        if (number >= positions[stops - 1])
            return scale.Colors[stops - 1];

        for (int i = 1; i < stops; i++)
        {
            if (number > positions[i])
                continue;

            double span = positions[i] - positions[i - 1];
            double fraction = span <= 0 ? 0 : (number - positions[i - 1]) / span;
            return Blend(scale.Colors[i - 1], scale.Colors[i], fraction);
        }

        return scale.Colors[stops - 1];
    }

    private ConditionalBar? Measure(ConditionalFormatRule rule, DataBar bar, in CellValue value)
    {
        if (!ConditionalRangeStatistics.TryNumber(value, DateSystemOf(), out double number))
            return null;

        ConditionalRangeStatistics statistics = Statistics(rule, occurrences: false);
        double low = Position(bar.Minimum, statistics, rule);
        double high = Position(bar.Maximum, statistics, rule);

        double position = Math.Clamp(high <= low ? 1 : (number - low) / (high - low), 0, 1);
        double minimumLength = Math.Clamp(bar.MinimumLength, 0, 100) / 100.0;
        double maximumLength = Math.Max(
            minimumLength,
            Math.Clamp(bar.MaximumLength, 0, 100) / 100.0);
        double fraction = minimumLength + (position * (maximumLength - minimumLength));
        return new ConditionalBar(fraction, bar.Color, bar.ShowValue);
    }

    private ConditionalIcon? Choose(ConditionalFormatRule rule, IconSet icons, in CellValue value)
    {
        if (icons.Values.Count == 0)
            return null;

        if (!ConditionalRangeStatistics.TryNumber(value, DateSystemOf(), out double number))
            return null;

        ConditionalRangeStatistics statistics = Statistics(rule, occurrences: false);

        // The bands are named by their lower bounds, lowest first; a value belongs to the
        // highest band whose bound it reaches.
        int band = 0;
        for (int i = 1; i < icons.Values.Count; i++)
        {
            ConditionalValue threshold = icons.Values[i];
            double boundary = Position(threshold, statistics, rule);
            if (threshold.GreaterThanOrEqual ? number >= boundary : number > boundary)
                band = i;
        }

        int index = icons.Reverse ? icons.Values.Count - 1 - band : band;
        return new ConditionalIcon(icons.Name, index, icons.Values.Count, icons.ShowValue);
    }

    /// <summary>Where one stop of a graded rule sits on the number line.</summary>
    private double Position(ConditionalValue stop, ConditionalRangeStatistics statistics, ConditionalFormatRule rule)
    {
        switch (stop.Type)
        {
            case ConditionalValueType.Minimum:
                return statistics.Minimum;
            case ConditionalValueType.Maximum:
                return statistics.Maximum;

            case ConditionalValueType.Percent:
            {
                double percent = StopParameter(stop.Value, rule, fallback: 0);
                return statistics.Minimum + ((statistics.Maximum - statistics.Minimum) * percent / 100);
            }

            case ConditionalValueType.Percentile:
                return statistics.Percentile(StopParameter(stop.Value, rule, fallback: 0));

            case ConditionalValueType.Number:
                return StopParameter(stop.Value, rule, fallback: 0);

            case ConditionalValueType.Formula:
                return StopParameter(stop.Value, rule, statistics.Minimum);

            default:
                return statistics.Minimum;
        }
    }

    private ConditionalRangeStatistics Statistics(ConditionalFormatRule rule, bool occurrences)
    {
        if (_statistics.TryGetValue(rule, out ConditionalRangeStatistics? cached))
            return cached;

        ConditionalRangeStatistics gathered =
            ConditionalRangeStatistics.Gather(_sheet, rule.Ranges, occurrences, DateSystemOf());

        _statistics[rule] = gathered;
        return gathered;
    }

    private ExcelDateSystem DateSystemOf() => _sheet.Workbook.DateSystem;

    /// <summary>
    /// The X parameter of a graded stop. BIFF and SpreadsheetML both allow the parameter of a
    /// number, percent, or percentile stop to come from a formula without changing the stop type.
    /// </summary>
    private double StopParameter(string? text, ConditionalFormatRule rule, double fallback)
    {
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            return value;
        if (string.IsNullOrWhiteSpace(text))
            return fallback;

        if (!CanEvaluateFormula)
        {
            _unsupported.Add("conditional-format-formula");
            return fallback;
        }

        string formula = text;
        if (_stopParameters.TryGetValue((rule, formula), out double? cached))
            return cached ?? fallback;

        CellRef anchor = rule.Ranges.Count == 0 ? new CellRef(1, 1) : rule.Ranges[0].Start;
        CellValue result = _evaluateRelative is { } evaluateRelative
            ? evaluateRelative(formula, anchor, anchor)
            : _evaluate!(formula, anchor);
        double? parameter = ConditionalRangeStatistics.TryNumber(result, DateSystemOf(), out double number)
            ? number
            : null;
        _stopParameters[(rule, formula)] = parameter;
        return parameter ?? fallback;
    }

    private static bool IsBlank(in CellValue value)
    {
        if (value.IsBlank)
            return true;
        if (value.Kind != CellValueKind.Text)
            return false;

        // The format defines this as LEN(TRIM(cell))=0. Excel's TRIM removes U+0020, not every
        // Unicode whitespace character, so tabs and non-breaking spaces remain nonblank.
        foreach (char character in value.GetText())
        {
            if (character != ' ')
                return false;
        }

        return true;
    }

    private static string Text(in CellValue value) =>
        value.Kind == CellValueKind.Text ? value.GetText() : value.ToString();

    /// <summary>
    /// Orders two values as Excel's comparisons do, or answers <see cref="int.MinValue"/> when
    /// they are not of comparable kinds — a number against a string, say, which no comparison
    /// rule paints.
    /// </summary>
    private int Compare(in CellValue left, in CellValue right)
    {
        ExcelDateSystem system = DateSystemOf();
        bool leftIsNumber = ConditionalRangeStatistics.TryNumber(left, system, out double leftNumber);
        bool rightIsNumber = ConditionalRangeStatistics.TryNumber(right, system, out double rightNumber);

        if (leftIsNumber && rightIsNumber)
            return leftNumber.CompareTo(rightNumber);

        if (left.Kind == CellValueKind.Text && right.Kind == CellValueKind.Text)
            return string.Compare(left.GetText(), right.GetText(), StringComparison.OrdinalIgnoreCase);

        // A blank cell compares as zero or as the empty string, which is what Excel shows.
        if (left.IsBlank && rightIsNumber)
            return 0d.CompareTo(rightNumber);
        if (left.IsBlank && right.Kind == CellValueKind.Text)
            return string.Compare("", right.GetText(), StringComparison.OrdinalIgnoreCase);

        return int.MinValue;
    }

    /// <summary>
    /// Adds a lower-priority format without replacing properties a higher-priority one named.
    /// </summary>
    private static DifferentialStyle Merge(DifferentialStyle? higher, DifferentialStyle lower)
    {
        if (higher is null)
            return lower;

        return new DifferentialStyle
        {
            Bold = higher.Bold ?? lower.Bold,
            Italic = higher.Italic ?? lower.Italic,
            Underline = higher.Underline ?? lower.Underline,
            Strikethrough = higher.Strikethrough ?? lower.Strikethrough,
            FontColor = higher.FontColor.IsAuto ? lower.FontColor : higher.FontColor,
            FillColor = higher.FillColor.IsAuto ? lower.FillColor : higher.FillColor,
            NumberFormat = higher.NumberFormat ?? lower.NumberFormat,
            Border = MergeBorders(higher.Border, lower.Border),
        };
    }

    private static BorderStyle? MergeBorders(BorderStyle? higher, BorderStyle? lower)
    {
        if (higher is null)
            return lower;
        if (lower is null)
            return higher;

        // BorderStyle predates differential formatting and has no per-edge presence bits. In a
        // dxf, however, edges are independent, so the useful convention is that Line=None means
        // "not supplied" while combining rules. It cannot express an explicit edge clear; that
        // remains a documented model limitation.
        bool higherDiagonal = higher.Diagonal.Line != BorderLineStyle.None;
        return new BorderStyle
        {
            Left = Pick(higher.Left, lower.Left),
            Right = Pick(higher.Right, lower.Right),
            Top = Pick(higher.Top, lower.Top),
            Bottom = Pick(higher.Bottom, lower.Bottom),
            Diagonal = Pick(higher.Diagonal, lower.Diagonal),
            DiagonalUp = higherDiagonal ? higher.DiagonalUp : lower.DiagonalUp,
            DiagonalDown = higherDiagonal ? higher.DiagonalDown : lower.DiagonalDown,
        };

        static BorderEdge Pick(BorderEdge first, BorderEdge second) =>
            first.Line == BorderLineStyle.None ? second : first;
    }

    /// <summary>Mixes two colours, channel by channel.</summary>
    private ExcelColor Blend(ExcelColor from, ExcelColor to, double fraction)
    {
        from = _sheet.Workbook.Theme.Resolve(from);
        to = _sheet.Workbook.Theme.Resolve(to);
        if (from.IsAuto || to.IsAuto)
            return fraction < 0.5 ? from : to;

        double amount = Math.Clamp(fraction, 0, 1);
        uint left = from.Argb;
        uint right = to.Argb;

        return ExcelColor.FromRgb(
            Channel(left >> 16, right >> 16, amount),
            Channel(left >> 8, right >> 8, amount),
            Channel(left, right, amount));

        static byte Channel(uint from, uint to, double amount) =>
            (byte)Math.Round(((from & 0xFF) * (1 - amount)) + ((to & 0xFF) * amount));
    }
}
