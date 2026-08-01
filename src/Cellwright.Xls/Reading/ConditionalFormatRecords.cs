using System.Buffers.Binary;
using Cellwright.Styles;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Collects the conditional formatting of one sheet. The rules come in groups: a
/// <c>CondFmt</c> record names the ranges, and the <c>CF</c> records after it are the rules
/// over those ranges. The newer <c>CondFmt12</c>/<c>CF12</c> pair is the same idea with room
/// for the graded rule types Excel 2007 added.
/// </summary>
internal sealed class ConditionalFormatCollector
{
    private readonly List<ConditionalFormatRule> _rules = [];

    /// <summary>Range lists by the group id that <c>CFEx</c> records point back at.</summary>
    private readonly Dictionary<int, IReadOnlyList<RangeRef>> _groups = [];

    private IReadOnlyList<RangeRef> _ranges = [];
    private int _priority = 1;
    private int _unmodelled;

    /// <summary>Colors are palette slots here, so the workbook's palette has to be at hand.</summary>
    private XlsGlobals? _globals;

    /// <summary>Gives the collector the palette to resolve color indexes against.</summary>
    public void Attach(XlsGlobals globals) => _globals = globals;

    /// <summary>Reads the group header, which carries the ranges the rules after it cover.</summary>
    /// <param name="body">Record body.</param>
    /// <param name="isVersion12">Whether this is the 2007 form, which adds a 12-byte prefix.</param>
    public void OnCondFmt(ReadOnlySpan<byte> body, bool isVersion12)
    {
        // CondFmt12 wraps the same structure in a future-record header.
        int offset = isVersion12 ? 12 : 0;
        if (offset + 12 > body.Length)
            return;

        // The rule count, then a flags word whose upper bits are the group's own id, then the
        // bounding box, and only then the real range list.
        int id = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]) >> 1;
        offset += 12;
        _ranges = ReadRanges(body, offset);
        _groups[id] = _ranges;
    }

    /// <summary>
    /// Follows a <c>CFEx</c> record, which says that the rules after it belong to a group
    /// stated earlier rather than to the one most recently read.
    /// </summary>
    public void OnCfEx(ReadOnlySpan<byte> body)
    {
        // Twelve bytes of future-record header, a flag saying which kind of rule is extended,
        // then the id of the group.
        const int IdOffset = 16;
        _ranges = IdOffset + 2 <= body.Length &&
            _groups.TryGetValue(BinaryPrimitives.ReadUInt16LittleEndian(body[IdOffset..]), out IReadOnlyList<RangeRef>? group)
                ? group
                : [];
    }

    /// <summary>Reads a classic rule: a comparison or a formula, with the format to apply.</summary>
    public void OnCf(ReadOnlySpan<byte> body, PtgDecoder decoder, int sheetIndex)
    {
        if (body.Length < 6)
            return;

        if (_ranges.Count == 0)
        {
            _unmodelled++;
            return;
        }

        byte conditionType = body[0];
        byte comparison = body[1];
        int cce1 = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        int cce2 = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);

        int offset = 6;
        DifferentialStyle style = DifferentialFormat.Read(body, ref offset, _globals);
        string?[] formulas = ReadFormulas(body, ref offset, cce1, cce2, decoder, sheetIndex);

        AddComparisonRule(conditionType, comparison, formulas, style, stopIfTrue: false);
    }

    /// <summary>
    /// Reads a 2007 rule. The comparison types are the same as the classic ones with a longer
    /// format block; the graded types — color scales, data bars, icon sets — are new, and their
    /// parameters trail everything else in the record.
    /// </summary>
    public void OnCf12(ReadOnlySpan<byte> body, PtgDecoder decoder, int sheetIndex)
    {
        // 12 bytes of future-record header, then the same opening fields as CF.
        const int Prefix = 12;
        if (body.Length < Prefix + 6)
            return;

        if (_ranges.Count == 0)
        {
            _unmodelled++;
            return;
        }

        byte conditionType = body[Prefix];
        byte comparison = body[Prefix + 1];
        int cce1 = BinaryPrimitives.ReadUInt16LittleEndian(body[(Prefix + 2)..]);
        int cce2 = BinaryPrimitives.ReadUInt16LittleEndian(body[(Prefix + 4)..]);

        int offset = Prefix + 6;
        DifferentialStyle style = ReadExtendedFormat(body, ref offset);
        string?[] formulas = ReadFormulas(body, ref offset, cce1, cce2, decoder, sheetIndex);

        // The activity formula, the flags, the priority and the template parameters sit between
        // the formulas and the graded parameters, all of them fixed width bar the first.
        SkipParsedFormula(body, ref offset);
        bool stopIfTrue = offset < body.Length && (body[offset] & 0x02) != 0;
        offset += 1;
        int priority = offset + 2 <= body.Length ? BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]) : 0;
        offset += 2 + 2;                    // Priority and the template it was created from.
        if (offset < body.Length)
            offset += 1 + body[offset];     // Template parameter block, a length and its bytes.

        ConditionalFormatRule? graded = conditionType switch
        {
            0x03 => ReadColorScale(body, offset),
            0x04 => ReadDataBar(body, offset),
            0x06 => ReadIconSet(body, offset),
            _ => null,
        };

        if (graded is not null)
        {
            graded.Priority = priority > 0 ? priority : _priority++;
            _rules.Add(graded);
            return;
        }

        if (conditionType is 0x01 or 0x02)
        {
            AddComparisonRule(conditionType, comparison, formulas, style, stopIfTrue, priority);
            return;
        }

        _unmodelled++;
    }

    /// <summary>Attaches the rules to the sheet, reporting the ones that had no model.</summary>
    /// <remarks>
    /// The rules are added as the model's own rather than as the file's. A workbook read from
    /// BIFF has no markup to replay on save, so the modern writer only emits rules it is told
    /// belong to the model.
    /// </remarks>
    public void Finish(Worksheet sheet, Action<WorkbookWarning> warn)
    {
        foreach (ConditionalFormatRule rule in _rules)
            sheet.ConditionalFormats.Add(rule);

        if (_unmodelled > 0)
        {
            warn(new WorkbookWarning(
                "CfRule12Skipped",
                $"'{sheet.Name}' has {_unmodelled} conditional formatting rules of a kind this library does not " +
                "model; they were left out."));
        }

        _rules.Clear();
        _groups.Clear();
        _ranges = [];
        _unmodelled = 0;
    }

    private void AddComparisonRule(
        byte conditionType,
        byte comparison,
        string?[] formulas,
        DifferentialStyle style,
        bool stopIfTrue,
        int priority = 0)
    {
        // A rule whose formula could not be decoded would highlight the wrong cells, which is
        // worse than not highlighting at all.
        if (conditionType == 0x02 && formulas[0] is null)
        {
            _unmodelled++;
            return;
        }

        List<string> texts = [];
        foreach (string? formula in formulas)
        {
            if (formula is not null)
                texts.Add(formula);
        }

        if (conditionType == 0x01 && texts.Count == 0)
        {
            _unmodelled++;
            return;
        }

        _rules.Add(new ConditionalFormatRule
        {
            Ranges = _ranges,
            Type = conditionType == 0x02 ? ConditionalFormatType.Expression : ConditionalFormatType.CellIs,
            Operator = conditionType == 0x02 ? ConditionalOperator.None : MapOperator(comparison),
            Formulas = texts,
            Style = style.IsEmpty ? null : style,
            StopIfTrue = stopIfTrue,
            Priority = priority > 0 ? priority : _priority++,
        });
    }

    /// <summary>Reads the two rule formulas, each anchored at the top-left cell of the first range.</summary>
    private string?[] ReadFormulas(
        ReadOnlySpan<byte> body, ref int offset, int cce1, int cce2, PtgDecoder decoder, int sheetIndex)
    {
        CellRef anchor = _ranges.Count > 0 ? _ranges[0].Start : new CellRef(1, 1);
        return
        [
            ReadFormula(body, ref offset, cce1, decoder, sheetIndex, anchor),
            ReadFormula(body, ref offset, cce2, decoder, sheetIndex, anchor),
        ];
    }

    private static string? ReadFormula(
        ReadOnlySpan<byte> body, ref int offset, int length, PtgDecoder decoder, int sheetIndex, CellRef anchor)
    {
        length = Math.Min(length, Math.Max(0, body.Length - offset));
        if (length <= 0)
            return null;

        string? text = decoder.Decode(body.Slice(offset, length), anchor, sheetIndex);
        offset += length;
        return text;
    }

    /// <summary>Steps over a formula that states its own length, which the graded rules use.</summary>
    private static void SkipParsedFormula(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 2 > body.Length)
        {
            offset = body.Length;
            return;
        }

        int length = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += 2 + Math.Min(length, Math.Max(0, body.Length - offset - 2));
    }

    /// <summary>Reads the format block of a 2007 rule, which wraps the classic one in a length.</summary>
    private DifferentialStyle ReadExtendedFormat(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return DifferentialStyle.Empty;

        int size = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;

        // Zero means no formatting at all, and then two reserved bytes stand in for the block.
        if (size == 0)
        {
            offset += 2;
            return DifferentialStyle.Empty;
        }

        int end = Math.Min(offset + size, body.Length);
        int cursor = offset;
        DifferentialStyle style = DifferentialFormat.Read(body[..end], ref cursor, _globals);
        offset = end;
        return style;
    }

    private ConditionalFormatRule? ReadColorScale(ReadOnlySpan<byte> body, int offset)
    {
        // Two counts of control points, then the value stops, then the colors.
        if (offset + 6 > body.Length)
            return null;

        int stops = body[offset + 3];
        if (stops is not (2 or 3))
            return null;

        int cursor = offset + 6;
        var values = new List<ConditionalValue>(stops);
        for (int i = 0; i < stops; i++)
        {
            if (!TryReadValueObject(body, ref cursor, out ConditionalValue value))
                return null;

            values.Add(value);
            cursor += 8;    // The position of this stop along the curve, which the model derives.
        }

        var colors = new List<ExcelColor>(stops);
        for (int i = 0; i < stops; i++)
        {
            if (cursor + 24 > body.Length)
                return null;

            colors.Add(ReadColor(body.Slice(cursor + 8, 16)));
            cursor += 24;   // The numeric position of the color, then the color itself.
        }

        return new ConditionalFormatRule
        {
            Ranges = _ranges,
            Type = ConditionalFormatType.ColorScale,
            ColorScale = new ColorScale { Values = values, Colors = colors },
        };
    }

    private ConditionalFormatRule? ReadDataBar(ReadOnlySpan<byte> body, int offset)
    {
        // Flags, the two length percentages, then the bar color and the two value stops.
        if (offset + 22 > body.Length)
            return null;

        bool showValue = (body[offset + 3] & 0x02) != 0;
        ExcelColor color = ReadColor(body.Slice(offset + 6, 16));

        int cursor = offset + 22;
        if (!TryReadValueObject(body, ref cursor, out ConditionalValue minimum) ||
            !TryReadValueObject(body, ref cursor, out ConditionalValue maximum))
        {
            return null;
        }

        return new ConditionalFormatRule
        {
            Ranges = _ranges,
            Type = ConditionalFormatType.DataBar,
            DataBar = new DataBar
            {
                Minimum = minimum,
                Maximum = maximum,
                Color = color,
                ShowValue = showValue,
            },
        };
    }

    private ConditionalFormatRule? ReadIconSet(ReadOnlySpan<byte> body, int offset)
    {
        if (offset + 6 > body.Length)
            return null;

        int states = body[offset + 3];
        if (states is < 3 or > 5)
            return null;

        string? name = IconSetName(body[offset + 4], states);
        if (name is null)
            return null;

        bool iconOnly = (body[offset + 5] & 0x01) != 0;
        bool reverse = (body[offset + 5] & 0x04) != 0;

        int cursor = offset + 6;
        var values = new List<ConditionalValue>(states);
        for (int i = 0; i < states; i++)
        {
            if (!TryReadValueObject(body, ref cursor, out ConditionalValue value))
                return null;

            values.Add(value);
            cursor += 5;    // Whether the threshold is inclusive, then four unused bytes.
        }

        return new ConditionalFormatRule
        {
            Ranges = _ranges,
            Type = ConditionalFormatType.IconSet,
            IconSet = new IconSet
            {
                Name = name,
                Values = values,
                ShowValue = !iconOnly,
                Reverse = reverse,
            },
        };
    }

    /// <summary>
    /// Reads one value object: how the stop is placed, and either the formula that computes it
    /// or the number that is it.
    /// </summary>
    private static bool TryReadValueObject(ReadOnlySpan<byte> body, ref int offset, out ConditionalValue value)
    {
        value = ConditionalValue.Minimum;
        if (offset + 3 > body.Length)
            return false;

        byte type = body[offset];
        int formulaLength = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 1)..]);
        offset += 3;

        if (formulaLength > 0)
        {
            // A stop given by a formula: the model keeps the expression, not its result, and
            // rendering it back is the writer's problem — which this library leaves to xlsx.
            offset += Math.Min(formulaLength, Math.Max(0, body.Length - offset));
            value = new ConditionalValue(ConditionalValueType.Formula, null);
            return type != 0x00;
        }

        // The minimum and the maximum of the range carry no number of their own.
        if (type is 0x02 or 0x03)
        {
            value = type == 0x02 ? ConditionalValue.Minimum : ConditionalValue.Maximum;
            return true;
        }

        if (offset + 8 > body.Length)
            return false;

        double number = BinaryPrimitives.ReadDoubleLittleEndian(body[offset..]);
        offset += 8;
        value = type switch
        {
            0x01 => ConditionalValue.Number(number),
            0x04 => ConditionalValue.Percent(number),
            0x05 => ConditionalValue.Percentile(number),
            _ => ConditionalValue.Number(number),
        };

        return true;
    }

    /// <summary>
    /// Reads a 16-byte color: how it is stated, the value, and a tint the model drops. The
    /// value is four bytes in painting order — red, green, blue, alpha — not a packed integer.
    /// </summary>
    private static ExcelColor ReadColor(ReadOnlySpan<byte> body)
    {
        uint kind = BinaryPrimitives.ReadUInt32LittleEndian(body);
        return kind switch
        {
            // An index into the palette, an explicit RGB, or a theme slot the model cannot name.
            0x01 => ExcelColor.FromIndexed(body[4] & 0x7F),
            0x02 => ExcelColor.FromRgb(body[4], body[5], body[6]),
            _ => ExcelColor.Auto,
        };
    }

    /// <summary>
    /// Names the icon set the way the modern format does, since the model speaks that. The
    /// binary set numbers are the same in BIFF8 and BIFF12, so the <c>.xlsb</c> reader shares
    /// this table.
    /// </summary>
    internal static string? IconSetName(byte id, int states) => id switch
    {
        0x00 => "3Arrows",
        0x01 => "3ArrowsGray",
        0x02 => "3Flags",
        0x03 => "3TrafficLights1",
        0x04 => "3TrafficLights2",
        0x05 => "3Signs",
        0x06 => "3Symbols",
        0x07 => "3Symbols2",
        0x08 => "4Arrows",
        0x09 => "4ArrowsGray",
        0x0A => "4RedToBlack",
        0x0B => "4Rating",
        0x0C => "4TrafficLights",
        0x0D => "5Arrows",
        0x0E => "5ArrowsGray",
        0x0F => "5Rating",
        0x10 => "5Quarters",
        _ => states switch
        {
            3 => "3TrafficLights1",
            4 => "4TrafficLights",
            5 => "5Rating",
            _ => null,
        },
    };

    /// <summary>Reads the range list that closes the group header.</summary>
    private static IReadOnlyList<RangeRef> ReadRanges(ReadOnlySpan<byte> body, int offset)
    {
        if (offset + 2 > body.Length)
            return [];

        int count = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += 2;

        var ranges = new List<RangeRef>(Math.Min(count, 1026));
        for (int i = 0; i < count && offset + 8 <= body.Length; i++, offset += 8)
        {
            int firstRow = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]) + 1;
            int lastRow = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 2)..]) + 1;
            int firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 4)..]) + 1;
            int lastColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[(offset + 6)..]) + 1;
            if (!SheetLimits.IsValidRow(lastRow) || !SheetLimits.IsValidColumn(lastColumn))
                continue;

            ranges.Add(new RangeRef(new CellRef(firstRow, firstColumn), new CellRef(lastRow, lastColumn)));
        }

        return ranges;
    }

    private static ConditionalOperator MapOperator(byte value) => value switch
    {
        0x01 => ConditionalOperator.Between,
        0x02 => ConditionalOperator.NotBetween,
        0x03 => ConditionalOperator.Equal,
        0x04 => ConditionalOperator.NotEqual,
        0x05 => ConditionalOperator.GreaterThan,
        0x06 => ConditionalOperator.LessThan,
        0x07 => ConditionalOperator.GreaterThanOrEqual,
        0x08 => ConditionalOperator.LessThanOrEqual,
        _ => ConditionalOperator.None,
    };
}
