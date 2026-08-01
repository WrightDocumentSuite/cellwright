using Cellwright.Styles;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes a sheet's conditional formatting as the classic BIFF8 records: a
/// <see cref="BiffRecordType.CondFmt"/> heading each group of rules, and a
/// <see cref="BiffRecordType.Cf"/> for each rule in it (MS-XLS 2.4.42 and 2.4.41).
/// </summary>
/// <remarks>
/// <para>
/// The classic records know two kinds of rule: a comparison of the cell against one or two
/// formulas, and a formula that is simply true or false. Everything Excel added later — colour
/// scales, data bars, icon sets, ranks, averages, date periods — lives in CF12 records that this
/// writer does not produce, and those rules are reported rather than approximated.
/// </para>
/// <para>
/// The text and blank and error rules are a middle case: they have no record of their own, but
/// each is exactly equivalent to a formula, and it is the very formula Excel itself stores in the
/// <c>formula</c> child of those rules in xlsx. Lowering them loses nothing a reader could see,
/// so they are written rather than dropped.
/// </para>
/// </remarks>
internal static class ConditionalFormatWriter
{
    /// <summary>Rules one CondFmt group may carry: the record's count field allows one to three.</summary>
    private const int MaxRulesPerGroup = 3;

    /// <summary>Ranges one group may name.</summary>
    private const int MaxRangesPerGroup = 1026;

    /// <summary>A comparison of the cell against one or two formulas.</summary>
    private const byte ConditionCompare = 0x01;

    /// <summary>A formula that decides the rule on its own.</summary>
    private const byte ConditionFormula = 0x02;

    /// <summary>Writes the sheet's rules, reporting the ones the format cannot carry.</summary>
    /// <param name="writer">Where to write.</param>
    /// <param name="sheet">The sheet whose rules to write.</param>
    /// <param name="encoder">The formula encoder for this sheet.</param>
    /// <param name="context">The write context, for the palette and the warnings.</param>
    public static void Write(BiffWriter writer, Worksheet sheet, PtgEncoder encoder, XlsWriteContext context)
    {
        if (sheet.ConditionalFormats.Count == 0)
            return;

        int skipped = 0;
        int identifier = 1;

        foreach (List<ConditionalFormatRule> group in Group(sheet, ref skipped, context))
        {
            // The record counts its rules in a field that allows three, so a longer group is
            // written as several groups over the same cells rather than truncated.
            for (int start = 0; start < group.Count; start += MaxRulesPerGroup)
            {
                List<ConditionalFormatRule> batch =
                    group.GetRange(start, Math.Min(MaxRulesPerGroup, group.Count - start));

                var bodies = new List<byte[]>(batch.Count);
                foreach (ConditionalFormatRule rule in batch)
                {
                    if (Encode(rule, encoder, context) is { } body)
                        bodies.Add(body);
                    else
                        skipped++;
                }

                if (bodies.Count == 0)
                    continue;

                WriteGroup(writer, batch[0].Ranges, bodies, identifier++);
            }
        }

        if (skipped > 0)
        {
            context.Warn(new WorkbookWarning(
                "ConditionalFormatsNotWritten",
                $"'{sheet.Name}' has {skipped} conditional formatting rule(s) of a kind the .xls format " +
                "keeps in records this library does not write — colour scales, data bars, icon sets, ranks, " +
                "averages, duplicates and date periods among them. They are not in the file; the rest are."));
        }
    }

    /// <summary>
    /// The sheet's rules gathered into groups that cover the same cells, in priority order.
    /// </summary>
    /// <remarks>
    /// A CondFmt names one set of cells for all the rules under it, so rules can only share a
    /// record when they cover the same cells. Grouping by the range set is therefore not an
    /// optimisation but the shape the format requires.
    /// </remarks>
    private static List<List<ConditionalFormatRule>> Group(
        Worksheet sheet, ref int skipped, XlsWriteContext context)
    {
        var byRanges = new Dictionary<string, List<ConditionalFormatRule>>(StringComparer.Ordinal);
        var order = new List<string>();

        var ordered = new List<ConditionalFormatRule>(sheet.ConditionalFormats);
        ordered.Sort(static (left, right) => left.Priority.CompareTo(right.Priority));

        foreach (ConditionalFormatRule rule in ordered)
        {
            if (rule.IsOpaque || rule.Ranges.Count == 0 || rule.Ranges.Count > MaxRangesPerGroup)
            {
                skipped++;
                continue;
            }

            string key = string.Join(' ', rule.Ranges);
            if (!byRanges.TryGetValue(key, out List<ConditionalFormatRule>? bucket))
            {
                bucket = [];
                byRanges[key] = bucket;
                order.Add(key);
            }

            bucket.Add(rule);
        }

        _ = context;
        var groups = new List<List<ConditionalFormatRule>>(order.Count);
        foreach (string key in order)
            groups.Add(byRanges[key]);

        return groups;
    }

    /// <summary>Writes one CondFmt heading and the CF records under it.</summary>
    private static void WriteGroup(
        BiffWriter writer, IReadOnlyList<RangeRef> ranges, List<byte[]> rules, int identifier)
    {
        using var body = new BiffWriter(64);
        body.AppendUInt16(rules.Count);

        // The identifier is what a CFEx record would use to name this group; the tough-recalc
        // bit is a redraw hint that costs nothing to leave clear.
        body.AppendUInt16(identifier & 0x7FFF);

        RangeRef bounds = Bounds(ranges);
        AppendRef(body, bounds);

        body.AppendUInt16(ranges.Count);
        foreach (RangeRef range in ranges)
            AppendRef(body, range);

        writer.WriteRecord(BiffRecordType.CondFmt, body.Written.Span);

        foreach (byte[] rule in rules)
            writer.WriteRecord(BiffRecordType.Cf, rule);
    }

    /// <summary>The smallest rectangle covering every range of a group.</summary>
    private static RangeRef Bounds(IReadOnlyList<RangeRef> ranges)
    {
        int firstRow = int.MaxValue;
        int lastRow = 0;
        int firstColumn = int.MaxValue;
        int lastColumn = 0;

        foreach (RangeRef range in ranges)
        {
            firstRow = Math.Min(firstRow, range.Start.Row);
            lastRow = Math.Max(lastRow, range.End.Row);
            firstColumn = Math.Min(firstColumn, range.Start.Column);
            lastColumn = Math.Max(lastColumn, range.End.Column);
        }

        return new RangeRef(new CellRef(firstRow, firstColumn), new CellRef(lastRow, lastColumn));
    }

    /// <summary>Writes a Ref8U: the rows and columns of a rectangle, each zero-based.</summary>
    private static void AppendRef(BiffWriter body, RangeRef range)
    {
        body.AppendUInt16(range.Start.Row - 1);
        body.AppendUInt16(range.End.Row - 1);
        body.AppendUInt16(range.Start.Column - 1);
        body.AppendUInt16(range.End.Column - 1);
    }

    /// <summary>
    /// One rule as a CF record body, or null when the format has nowhere to put it.
    /// </summary>
    private static byte[]? Encode(ConditionalFormatRule rule, PtgEncoder encoder, XlsWriteContext context)
    {
        // A rule's relative references read as they do in the top-left cell of its first range.
        CellRef anchor = rule.Ranges[0].Start;

        byte condition;
        byte comparison;
        byte[]? first;
        byte[]? second = null;

        if (rule.Type == ConditionalFormatType.CellIs)
        {
            condition = ConditionCompare;
            comparison = (byte)rule.Operator;
            if (comparison is < 0x01 or > 0x08 || rule.Formulas.Count == 0)
                return null;

            first = encoder.Encode(rule.Formulas[0], anchor);
            if (first is null)
                return null;

            // Only Between and NotBetween take a second operand; the record requires the others
            // to leave it empty.
            if (comparison is 0x01 or 0x02)
            {
                if (rule.Formulas.Count < 2)
                    return null;

                second = encoder.Encode(rule.Formulas[1], anchor);
                if (second is null)
                    return null;
            }
        }
        else if (AsFormula(rule, anchor) is { } expression)
        {
            condition = ConditionFormula;
            comparison = 0;
            first = encoder.Encode(expression, anchor);
            if (first is null)
                return null;
        }
        else
        {
            return null;
        }

        using var body = new BiffWriter(160);
        body.AppendByte(condition);
        body.AppendByte(comparison);
        body.AppendUInt16(first.Length);
        body.AppendUInt16(second?.Length ?? 0);

        DifferentialFormatWriter.Write(body, rule.Style, context);

        body.Append(first);
        if (second is not null)
            body.Append(second);

        return body.ToArray();
    }

    /// <summary>
    /// The formula a rule is equivalent to, for the kinds the classic records have no condition
    /// for but that are expressible all the same. Null when the rule has no equivalent.
    /// </summary>
    /// <remarks>
    /// These are the formulas Excel itself writes into the <c>formula</c> child of the same rules
    /// in xlsx, so a file written this way says what the original said and is read back the same
    /// way by anything that reads the classic records.
    /// </remarks>
    private static string? AsFormula(ConditionalFormatRule rule, CellRef anchor)
    {
        string cell = $"{ColumnName.Format(anchor.Column)}{anchor.Row}";

        return rule.Type switch
        {
            ConditionalFormatType.Expression => rule.Formulas.Count > 0 ? rule.Formulas[0] : null,

            ConditionalFormatType.ContainsText when rule.Text is { Length: > 0 } =>
                $"NOT(ISERROR(SEARCH({Quote(rule.Text)},{cell})))",
            ConditionalFormatType.NotContainsText when rule.Text is { Length: > 0 } =>
                $"ISERROR(SEARCH({Quote(rule.Text)},{cell}))",
            ConditionalFormatType.BeginsWith when rule.Text is { Length: > 0 } =>
                $"LEFT({cell},{rule.Text.Length})={Quote(rule.Text)}",
            ConditionalFormatType.EndsWith when rule.Text is { Length: > 0 } =>
                $"RIGHT({cell},{rule.Text.Length})={Quote(rule.Text)}",

            ConditionalFormatType.ContainsBlanks => $"LEN(TRIM({cell}))=0",
            ConditionalFormatType.NotContainsBlanks => $"LEN(TRIM({cell}))>0",
            ConditionalFormatType.ContainsErrors => $"ISERROR({cell})",
            ConditionalFormatType.NotContainsErrors => $"NOT(ISERROR({cell}))",

            _ => null,
        };
    }

    private static string Quote(string text) =>
        $"\"{text.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
}
