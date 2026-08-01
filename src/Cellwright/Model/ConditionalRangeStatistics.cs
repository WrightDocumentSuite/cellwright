namespace Cellwright;

/// <summary>
/// What a conditional formatting rule needs to know about the range it covers, taken once.
/// </summary>
/// <remarks>
/// Most rule types cannot judge a cell on its own: the highest ten values, the ones above the
/// average, the duplicates and every graded rule are all statements about the range as a whole.
/// Gathering that per cell would walk the range once per cell; gathering it once per rule walks
/// it once, which is the difference between quadratic and linear on a column of a thousand.
/// </remarks>
internal sealed class ConditionalRangeStatistics
{
    private readonly double[] _sorted;
    private readonly Dictionary<string, int>? _occurrences;

    private ConditionalRangeStatistics(double[] sorted, Dictionary<string, int>? occurrences, double sum, double sumOfSquares)
    {
        _sorted = sorted;
        _occurrences = occurrences;
        Sum = sum;
        SumOfSquares = sumOfSquares;
    }

    /// <summary>How many numbers the range holds.</summary>
    public int Count => _sorted.Length;

    /// <summary>The lowest number, or zero when there are none.</summary>
    public double Minimum => _sorted.Length == 0 ? 0 : _sorted[0];

    /// <summary>The highest number, or zero when there are none.</summary>
    public double Maximum => _sorted.Length == 0 ? 0 : _sorted[^1];

    /// <summary>The total of the numbers.</summary>
    public double Sum { get; }

    /// <summary>The total of their squares, which is what the standard deviation is built from.</summary>
    public double SumOfSquares { get; }

    /// <summary>The mean, or zero when the range holds no numbers.</summary>
    public double Average => _sorted.Length == 0 ? 0 : Sum / _sorted.Length;

    /// <summary>The population standard deviation, as Excel's above-average rules use it.</summary>
    public double StandardDeviation
    {
        get
        {
            if (_sorted.Length == 0)
                return 0;

            double variance = (SumOfSquares / _sorted.Length) - (Average * Average);
            return variance <= 0 ? 0 : Math.Sqrt(variance);
        }
    }

    /// <summary>
    /// Gathers the statistics of every cell the rule covers.
    /// </summary>
    /// <param name="sheet">The sheet the ranges belong to.</param>
    /// <param name="ranges">The ranges the rule covers.</param>
    /// <param name="countOccurrences">Whether to count repeated values, which only two rule types need.</param>
    /// <param name="system">The workbook's date system, which fixes what a date's serial is.</param>
    public static ConditionalRangeStatistics Gather(
        Worksheet sheet, IReadOnlyList<RangeRef> ranges, bool countOccurrences, ExcelDateSystem system)
    {
        var numbers = new List<double>();
        Dictionary<string, int>? occurrences = countOccurrences
            ? new Dictionary<string, int>(StringComparer.Ordinal)
            : null;

        double sum = 0;
        double sumOfSquares = 0;

        foreach (RangeRef range in ranges)
        {
            for (int row = range.Start.Row; row <= range.End.Row; row++)
            {
                for (int column = range.Start.Column; column <= range.End.Column; column++)
                {
                    CellValue value = sheet.Values.Get(row, column);

                    if (occurrences is not null && !value.IsBlank)
                    {
                        string key = OccurrenceKey(value);
                        occurrences[key] = occurrences.GetValueOrDefault(key) + 1;
                    }

                    if (!TryNumber(value, system, out double number))
                        continue;

                    numbers.Add(number);
                    sum += number;
                    sumOfSquares += number * number;
                }
            }
        }

        double[] sorted = [.. numbers];
        Array.Sort(sorted);
        return new ConditionalRangeStatistics(sorted, occurrences, sum, sumOfSquares);
    }

    /// <summary>How many times a value appears in the range.</summary>
    /// <param name="value">The value to count.</param>
    public int OccurrencesOf(in CellValue value) =>
        value.IsBlank || _occurrences is null ? 0 : _occurrences.GetValueOrDefault(OccurrenceKey(value));

    /// <summary>
    /// The value at a percentile, interpolating between neighbours as Excel's PERCENTILE does.
    /// </summary>
    /// <param name="percent">Where in the range to look, 0 through 100.</param>
    public double Percentile(double percent)
    {
        if (_sorted.Length == 0)
            return 0;
        if (_sorted.Length == 1)
            return _sorted[0];

        double position = Math.Clamp(percent, 0, 100) / 100 * (_sorted.Length - 1);
        int lower = (int)Math.Floor(position);
        int upper = Math.Min(lower + 1, _sorted.Length - 1);
        return _sorted[lower] + ((_sorted[upper] - _sorted[lower]) * (position - lower));
    }

    /// <summary>
    /// The cut-off a top-or-bottom rule paints from: the value a cell has to reach to be among
    /// the highest, or stay under to be among the lowest.
    /// </summary>
    /// <param name="rank">How many values, or what share of them when <paramref name="percent"/>.</param>
    /// <param name="bottom">Whether the lowest are wanted rather than the highest.</param>
    /// <param name="percent">Whether the rank is a percentage rather than a count.</param>
    public double RankThreshold(int rank, bool bottom, bool percent)
    {
        if (_sorted.Length == 0)
            return 0;

        int wanted = percent
            ? (int)Math.Floor(_sorted.Length * Math.Clamp(rank, 0, 100) / 100.0)
            : rank;

        wanted = Math.Clamp(wanted, 1, _sorted.Length);
        return bottom ? _sorted[wanted - 1] : _sorted[^wanted];
    }

    /// <summary>
    /// A value as the number a rule compares against. A date counts as its serial number, which
    /// is what makes a colour scale over a column of dates grade them by when they are. A boolean
    /// does not: Excel's graded rules pass over them, and so do its averages.
    /// </summary>
    /// <param name="value">The cell's value.</param>
    /// <param name="system">The workbook's date system, which fixes what a date's serial is.</param>
    /// <param name="number">The number it counts as.</param>
    public static bool TryNumber(in CellValue value, ExcelDateSystem system, out double number)
    {
        switch (value.Kind)
        {
            case CellValueKind.Number:
                number = value.GetNumber();
                return true;

            case CellValueKind.DateTime:
            case CellValueKind.TimeSpan:
                number = value.ToSerialNumber(system);
                return true;

            default:
                number = 0;
                return false;
        }
    }

    /// <summary>
    /// How a value is identified when counting repeats. Excel compares what it shows rather than
    /// how it is stored, so the number 1 and the text "1" are one value to a duplicate rule.
    /// </summary>
    private static string OccurrenceKey(in CellValue value) => value.Kind switch
    {
        CellValueKind.Number => value.GetNumber().ToString("R", System.Globalization.CultureInfo.InvariantCulture),
        CellValueKind.Text => value.GetText(),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        _ => string.Empty,
    };
}
