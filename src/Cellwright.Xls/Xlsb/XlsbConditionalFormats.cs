using System.Buffers.Binary;
using Cellwright.Model;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Collects the conditional formatting of one sheet: a range, then the rules over it
/// (MS-XLSB 2.4.34 and 2.4.23).
/// </summary>
/// <remarks>
/// The binary package states a rule with two numbers where the older format used one. A rule
/// has a <em>type</em> saying broadly what it does and a <em>template</em> saying which of the
/// many conditions of that type it is, and the pair together is what the model calls a kind.
/// Colour scales, data bars and icon sets announce themselves in the type and then describe
/// themselves in the records that follow — <c>BrtBeginColorScale</c> and its kin, each holding
/// value stops (<c>BrtCFVO</c>) and colors (<c>BrtColor</c>) — which are collected until
/// <c>BrtEndCFRule</c> closes the rule.
/// </remarks>
internal sealed class XlsbConditionalFormats(PtgDecoder decoder)
{
    /// <summary>Fixed part of <c>BrtBeginCFRule</c> before its strings and formulas.</summary>
    private const int RuleHeader = 44;

    private readonly List<ConditionalFormatRule> _rules = [];
    private IReadOnlyList<RangeRef> _ranges = [];
    private int _unmodelled;

    // A graded rule under construction: its kind arrives in BrtBeginCFRule, its detail in the
    // child records, and only BrtEndCFRule tells whether enough of it showed up.
    private ConditionalFormatType? _pendingKind;
    private int _pendingPriority;
    private string? _pendingActivityFormula;
    private readonly List<ConditionalValue> _stops = [];
    private readonly List<ExcelColor> _colors = [];
    private bool _collecting;
    private bool _pendingInvalid;
    private bool _barShowValue;
    private int _barMinimumLength = 10;
    private int _barMaximumLength = 90;
    private int _iconSetId;
    private bool _iconOnly;
    private bool _iconReverse;

    /// <summary>Reads the header that opens a group: the ranges its rules apply to.</summary>
    public void BeginGroup(ReadOnlySpan<byte> body)
    {
        // ccf and fPivot, then the ranges the group covers.
        _ranges = body.Length >= 8 ? ReadRanges(body[8..]) : [];
    }

    /// <summary>Reads one rule of the group the reader is inside.</summary>
    public void ReadRule(ReadOnlySpan<byte> body)
    {
        if (body.Length < RuleHeader || _ranges.Count == 0)
            return;

        int type = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        int template = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]);
        int styleId = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[8..]);
        int priority = BinaryPrimitives.ReadInt32LittleEndian(body[12..]);
        int parameter = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[16..]);
        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body[28..]);

        int offset = RuleHeader;
        string? text = Biff12Values.ReadNullableString(body, ref offset);
        string? first = ReadFormula(body, ref offset, BinaryPrimitives.ReadUInt32LittleEndian(body[32..]));
        string? second = ReadFormula(body, ref offset, BinaryPrimitives.ReadUInt32LittleEndian(body[36..]));
        uint thirdPresent = BinaryPrimitives.ReadUInt32LittleEndian(body[40..]);
        string? third = ReadFormula(body, ref offset, thirdPresent);

        // The graded kinds (CFType 3, 4 and 6) describe themselves in the records that follow.
        ConditionalFormatType? graded = type switch
        {
            3 => ConditionalFormatType.ColorScale,
            4 => ConditionalFormatType.DataBar,
            6 => ConditionalFormatType.IconSet,
            _ => null,
        };
        if (graded is { } gradedKind)
        {
            if (thirdPresent != 0 && third is null)
            {
                _unmodelled++;
                return;
            }

            _pendingKind = gradedKind;
            _pendingPriority = priority > 0 ? priority : 0;
            _pendingActivityFormula = third;
            _pendingInvalid = false;
            // fStopTrue MUST be zero for graded rules. Do not propagate a malformed producer's
            // bit into the model, where it would change the evaluation of lower-priority rules.
            return;
        }

        if (Classify(type, template) is not { } kind)
        {
            _unmodelled++;
            return;
        }

        var formulas = new List<string>(2);
        if (first is not null)
            formulas.Add(first);
        if (second is not null)
            formulas.Add(second);

        _rules.Add(new ConditionalFormatRule
        {
            Ranges = _ranges,
            Type = kind,
            Operator = kind == ConditionalFormatType.CellIs ? MapOperator(parameter) : ConditionalOperator.None,
            Formulas = formulas,
            Text = text,
            Rank = kind is ConditionalFormatType.Top10 ? parameter : 0,
            Bottom = (flags & 0x00000008) != 0,
            Percent = (flags & 0x00000010) != 0,
            StopIfTrue = (flags & 0x00000002) != 0,
            Priority = priority > 0 ? priority : 0,
            SourceStyleId = styleId == -1 ? null : styleId,
        });
    }

    /// <summary>Opens the color scale block of the pending rule (MS-XLSB 2.4.28).</summary>
    public void BeginColorScale() =>
        _collecting = _pendingKind == ConditionalFormatType.ColorScale;

    /// <summary>Opens the data bar block: two length percentages, then the show flag (2.4.43).</summary>
    public void BeginDataBar(ReadOnlySpan<byte> body)
    {
        _collecting = _pendingKind == ConditionalFormatType.DataBar;
        if (!_collecting || body.Length < 3)
        {
            _collecting = false;
            _pendingInvalid = true;
            return;
        }

        _barMinimumLength = body[0];
        _barMaximumLength = body[1];
        _barShowValue = body[2] != 0;
        if (_barMinimumLength > _barMaximumLength || _barMaximumLength > 100)
        {
            _collecting = false;
            _pendingInvalid = true;
        }
    }

    /// <summary>Opens the icon set block: the set number and the display bits (2.4.91).</summary>
    public void BeginIconSet(ReadOnlySpan<byte> body)
    {
        _collecting = _pendingKind == ConditionalFormatType.IconSet;
        if (body.Length < 6)
        {
            _pendingInvalid = true;
            return;
        }

        _iconSetId = BinaryPrimitives.ReadInt32LittleEndian(body);
        int bits = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        _iconOnly = (bits & 0x02) != 0;
        // BrtBeginIconSet spells this flag backwards: zero means reversed order, while
        // one means the icon set's defined order. The model and OOXML both use the direct
        // meaning, so invert the wire bit here.
        _iconReverse = (bits & 0x04) == 0;
    }

    /// <summary>
    /// Reads one value stop (<c>BrtCFVO</c>, 2.4.334): how the threshold is determined, and
    /// either the number that is it or the formula that computes it.
    /// </summary>
    public void ReadValueObject(ReadOnlySpan<byte> body)
    {
        if (!_collecting)
            return;
        if (body.Length < 24)
        {
            _pendingInvalid = true;
            return;
        }

        int type = BinaryPrimitives.ReadInt32LittleEndian(body);
        if (type is not (1 or 2 or 3 or 4 or 5 or 7))
        {
            _pendingInvalid = true;
            return;
        }

        double number = BinaryPrimitives.ReadDoubleLittleEndian(body[4..]);
        bool isIcon = _pendingKind == ConditionalFormatType.IconSet;
        uint saveGreaterThanOrEqual = BinaryPrimitives.ReadUInt32LittleEndian(body[12..]);
        uint greaterThanOrEqualFlag = BinaryPrimitives.ReadUInt32LittleEndian(body[16..]);
        if (isIcon && (saveGreaterThanOrEqual != 1 || greaterThanOrEqualFlag > 1))
        {
            _pendingInvalid = true;
            return;
        }

        bool greaterThanOrEqual = !isIcon || greaterThanOrEqualFlag != 0;
        uint formulaLength = BinaryPrimitives.ReadUInt32LittleEndian(body[20..]);

        if (type is 2 or 3)
        {
            if (isIcon)
            {
                _pendingInvalid = true;
                return;
            }

            ConditionalValue boundary = type == 2 ? ConditionalValue.Minimum : ConditionalValue.Maximum;
            _stops.Add(boundary with { GreaterThanOrEqual = greaterThanOrEqual });
            return;
        }

        string? formula = null;
        if (formulaLength > 0)
        {
            // formula is a CFVOParsedFormula: cbFmla counts rgce, while the structure repeats
            // that count as cce before the tokens and carries an rgcb length after them.
            if (body.Length < 28)
            {
                _pendingInvalid = true;
                return;
            }

            uint tokenLength = BinaryPrimitives.ReadUInt32LittleEndian(body[24..]);
            if (tokenLength != formulaLength || tokenLength > 16_384 || tokenLength > body.Length - 28)
            {
                _pendingInvalid = true;
                return;
            }

            int tokenCount = (int)tokenLength;
            int extraLengthOffset = 28 + tokenCount;
            if (extraLengthOffset + 4 > body.Length)
            {
                _pendingInvalid = true;
                return;
            }

            uint extraLength = BinaryPrimitives.ReadUInt32LittleEndian(body[extraLengthOffset..]);
            if (extraLength > body.Length - extraLengthOffset - 4)
            {
                _pendingInvalid = true;
                return;
            }

            formula = decoder.Decode(body.Slice(28, tokenCount), _ranges[0].Start, sheetIndex: 0);
            if (formula is null)
            {
                _pendingInvalid = true;
                return;
            }
        }
        else if (type == 7)
        {
            _pendingInvalid = true;
            return;
        }

        if (formula is null && (!double.IsFinite(number) ||
            (type is 4 or 5 && (number < 0 || number > 100))))
        {
            _pendingInvalid = true;
            return;
        }

        ConditionalValue stop = (type, formula) switch
        {
            (1, string text) => new ConditionalValue(ConditionalValueType.Number, text),
            (4, string text) => new ConditionalValue(ConditionalValueType.Percent, text),
            (5, string text) => new ConditionalValue(ConditionalValueType.Percentile, text),
            (7, string text) => ConditionalValue.Formula(text),
            (4, _) => ConditionalValue.Percent(number),
            (5, _) => ConditionalValue.Percentile(number),
            _ => ConditionalValue.Number(number),
        };
        _stops.Add(stop with { GreaterThanOrEqual = greaterThanOrEqual });
    }

    /// <summary>Reads one color of the open block (<c>BrtColor</c>, 2.4.337).</summary>
    public void ReadColor(ReadOnlySpan<byte> body)
    {
        if (!_collecting)
            return;
        if (body.Length < 8)
        {
            _pendingInvalid = true;
            return;
        }

        _colors.Add(XlsbStyles.ReadColor(body));
    }

    /// <summary>
    /// Closes the current rule. A graded rule whose detail records added up becomes a model
    /// rule; one that did not stays honestly unread.
    /// </summary>
    public void EndRule()
    {
        if (_pendingKind is { } kind)
        {
            ConditionalFormatRule? rule = _pendingInvalid ? null : BuildGraded(kind);
            if (rule is null)
                _unmodelled++;
            else
                _rules.Add(rule);
        }

        _pendingKind = null;
        _pendingActivityFormula = null;
        _pendingInvalid = false;
        _collecting = false;
        _stops.Clear();
        _colors.Clear();
    }

    private ConditionalFormatRule? BuildGraded(ConditionalFormatType kind) => kind switch
    {
        ConditionalFormatType.ColorScale when _stops.Count >= 2 && _colors.Count == _stops.Count =>
            new ConditionalFormatRule
            {
                Ranges = _ranges,
                Type = kind,
                Priority = _pendingPriority,
                Formulas = ActivityFormulas(),
                ColorScale = new ColorScale { Values = [.. _stops], Colors = [.. _colors] },
            },
        ConditionalFormatType.DataBar when _stops.Count >= 2 =>
            new ConditionalFormatRule
            {
                Ranges = _ranges,
                Type = kind,
                Priority = _pendingPriority,
                Formulas = ActivityFormulas(),
                DataBar = new DataBar
                {
                    Minimum = _stops[0],
                    Maximum = _stops[1],
                    Color = _colors.Count > 0 ? _colors[0] : ExcelColor.FromRgb(0x63, 0x8E, 0xC6),
                    ShowValue = _barShowValue,
                    MinimumLength = _barMinimumLength,
                    MaximumLength = _barMaximumLength,
                },
            },
        ConditionalFormatType.IconSet when _stops.Count >= 3 && (uint)_iconSetId <= 0x10 &&
            Reading.ConditionalFormatCollector.IconSetName((byte)_iconSetId, _stops.Count) is { } name =>
            new ConditionalFormatRule
            {
                Ranges = _ranges,
                Type = kind,
                Priority = _pendingPriority,
                Formulas = ActivityFormulas(),
                IconSet = new IconSet
                {
                    Name = name,
                    Values = [.. _stops],
                    ShowValue = !_iconOnly,
                    Reverse = _iconReverse,
                },
            },
        _ => null,
    };

    private IReadOnlyList<string> ActivityFormulas() =>
        _pendingActivityFormula is null ? [] : [_pendingActivityFormula];

    /// <summary>Hands the rules to the sheet, once the whole part has been read.</summary>
    public void Finish(Worksheet sheet, Action<WorkbookWarning> warn)
    {
        foreach (ConditionalFormatRule rule in _rules)
            sheet.ConditionalFormats.AddFromFile(rule);

        if (_unmodelled > 0)
        {
            warn(new WorkbookWarning(
                "CfRule12Skipped",
                $"'{sheet.Name}' has {_unmodelled} conditional formatting rules of a kind this library does not " +
                "model; they were left out."));
        }

        _rules.Clear();
        _unmodelled = 0;
    }

    /// <summary>
    /// Maps the type and template pair onto what the model calls a rule kind. The ones with no
    /// entry here are the scales, bars and icon sets, whose detail lives in further records.
    /// </summary>
    private static ConditionalFormatType? Classify(int type, int template) => (type, template) switch
    {
        (1, _) => ConditionalFormatType.CellIs,
        (2, 0x01) => ConditionalFormatType.Expression,
        (2, 0x07) => ConditionalFormatType.UniqueValues,
        (2, 0x08) => ConditionalFormatType.ContainsText,
        (2, 0x09) => ConditionalFormatType.ContainsBlanks,
        (2, 0x0A) => ConditionalFormatType.NotContainsBlanks,
        (2, 0x0B) => ConditionalFormatType.ContainsErrors,
        (2, 0x0C) => ConditionalFormatType.NotContainsErrors,
        (2, 0x19 or 0x1A) => ConditionalFormatType.AboveAverage,
        (2, 0x1B) => ConditionalFormatType.DuplicateValues,
        (5, _) => ConditionalFormatType.Top10,
        _ => null,
    };

    /// <summary>The comparison of a cell-value rule, which arrives in the parameter field.</summary>
    private static ConditionalOperator MapOperator(int value) => value switch
    {
        1 => ConditionalOperator.Between,
        2 => ConditionalOperator.NotBetween,
        3 => ConditionalOperator.Equal,
        4 => ConditionalOperator.NotEqual,
        5 => ConditionalOperator.GreaterThan,
        6 => ConditionalOperator.LessThan,
        7 => ConditionalOperator.GreaterThanOrEqual,
        8 => ConditionalOperator.LessThanOrEqual,
        _ => ConditionalOperator.None,
    };

    /// <summary>
    /// A <c>CFParsedFormula</c>, present only when the length field before the strings said so.
    /// The header value is only a presence flag; cce and cb inside the structure delimit its
    /// token and ancillary-data streams.
    /// </summary>
    private string? ReadFormula(ReadOnlySpan<byte> body, ref int offset, uint present)
    {
        if (present == 0 || offset + 4 > body.Length)
            return null;

        uint tokenLength = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (tokenLength is 0 or > 16_384 || tokenLength > body.Length - offset)
        {
            offset = body.Length;
            return null;
        }

        int tokens = (int)tokenLength;
        ReadOnlySpan<byte> rgce = body.Slice(offset, tokens);
        offset += tokens;
        if (offset + 4 > body.Length)
        {
            offset = body.Length;
            return null;
        }

        uint extraLength = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (extraLength > body.Length - offset)
        {
            offset = body.Length;
            return null;
        }

        offset += (int)extraLength;
        // A rule's formula reads as it would in the top-left cell of the range it covers.
        CellRef anchor = _ranges.Count > 0 ? _ranges[0].Start : new CellRef(1, 1);
        return decoder.Decode(rgce, anchor, sheetIndex: 0);
    }

    private static IReadOnlyList<RangeRef> ReadRanges(ReadOnlySpan<byte> body)
    {
        if (body.Length < 4)
            return [];

        int count = BinaryPrimitives.ReadInt32LittleEndian(body);
        if (count is < 0 or > 8192)
            return [];

        var ranges = new List<RangeRef>(Math.Min(count, 64));
        for (int i = 0; i < count && 4 + ((i + 1) * 16) <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(4 + (i * 16), 16);
            ranges.Add(new RangeRef(
                new CellRef(
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(entry) + 1,
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(entry[8..]) + 1),
                new CellRef(
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(entry[4..]) + 1,
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(entry[12..]) + 1)));
        }

        return ranges;
    }
}
