using System.Buffers.Binary;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The fields added to BIFF8 conditional formatting by CF12. These byte-level fixtures keep
/// the assertions independent of whichever external spreadsheet corpus happens to be present.
/// </summary>
public sealed class Cf12FeatureTests
{
    [Fact]
    public void ADataBarKeepsItsLengthLimits()
    {
        byte[] parameters = new byte[28];
        parameters[3] = 0x02;               // fShowValue.
        parameters[4] = 15;                 // iPercentMin.
        parameters[5] = 80;                 // iPercentMax.
        BinaryPrimitives.WriteUInt32LittleEndian(parameters.AsSpan(6), 0x02); // Explicit RGB.
        parameters[10] = 0x63;
        parameters[11] = 0x8E;
        parameters[12] = 0xC6;
        parameters[13] = 0x80;
        parameters[22] = 0x02;              // CFVO minimum, with an empty formula.
        parameters[25] = 0x03;              // CFVO maximum, with an empty formula.

        ConditionalFormatRule rule = ReadRule(0x04, parameters);

        Assert.NotNull(rule.DataBar);
        Assert.Equal(15, rule.DataBar!.MinimumLength);
        Assert.Equal(80, rule.DataBar.MaximumLength);
        Assert.True(rule.DataBar.ShowValue);
        Assert.Equal(ExcelColor.FromArgb(0x80638EC6), rule.DataBar.Color);
    }

    [Theory]
    [InlineData(101, 101)]
    [InlineData(80, 20)]
    [InlineData(50, 50)]
    public void AnInvalidDataBarLengthIsReportedAndSkipped(byte minimum, byte maximum)
    {
        byte[] parameters = new byte[28];
        parameters[4] = minimum;
        parameters[5] = maximum;
        parameters[22] = 0x02;
        parameters[25] = 0x03;

        var globals = new XlsGlobals { Version = BiffVersion.Biff8 };
        var collector = new ConditionalFormatCollector();
        collector.Attach(globals);
        collector.OnCondFmt(Group(), isVersion12: true);
        collector.OnCf12(
            Rule(0x04, parameters, activityFormula: null, flags: 0),
            new PtgDecoder(globals, ["Sheet1"]),
            sheetIndex: 0);

        var warnings = new List<WorkbookWarning>();
        Worksheet sheet = Workbook.Create().Sheets.Add("Sheet1");
        collector.Finish(sheet, warnings.Add);

        Assert.Empty(sheet.ConditionalFormats);
        Assert.Contains(warnings, static warning => warning.Code == "CfRule12Skipped");
    }

    [Fact]
    public void AnIconSetKeepsWhetherEachThresholdIsInclusive()
    {
        byte[] parameters = new byte[54];
        parameters[3] = 3;                  // Three thresholds/states.
        parameters[4] = 3;                  // 3TrafficLights1.

        int cursor = 6;
        WriteThreshold(parameters, ref cursor, 0, inclusive: true);
        WriteThreshold(parameters, ref cursor, 33, inclusive: false);
        WriteThreshold(parameters, ref cursor, 67, inclusive: true);

        ConditionalFormatRule rule = ReadRule(0x06, parameters);

        Assert.NotNull(rule.IconSet);
        Assert.Equal(
            [true, false, true],
            rule.IconSet!.Values.Select(static value => value.GreaterThanOrEqual));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(2, 3)]
    [InlineData(3, 2)]
    public void AColorScaleKeepsThemeSlotsAndTints(int binaryTheme, int modelTheme)
    {
        byte[] parameters = new byte[76];
        parameters[3] = 2;                  // Two interpolation/control points.

        int cursor = 6;
        parameters[cursor] = 0x02;          // CFVO minimum, with an empty formula.
        cursor += 3 + 8;                    // CFVO plus numDomain.
        parameters[cursor] = 0x03;          // CFVO maximum, with an empty formula.
        cursor += 3 + 8;

        WriteThemeColor(parameters.AsSpan(cursor + 8), binaryTheme, tint: -0.25);
        cursor += 24;                       // numDomain plus CFColor.
        WriteThemeColor(parameters.AsSpan(cursor + 8), binaryTheme, tint: 0.5);

        ConditionalFormatRule rule = ReadRule(0x03, parameters);

        Assert.NotNull(rule.ColorScale);
        Assert.Equal(
            [ExcelColor.FromTheme(modelTheme, -0.25), ExcelColor.FromTheme(modelTheme, 0.5)],
            rule.ColorScale!.Colors);
    }

    [Theory]
    [InlineData(-1, 0.0)]
    [InlineData(11, 0.0)]
    [InlineData(0, 1.01)]
    [InlineData(0, -1.01)]
    [InlineData(0, double.NaN)]
    public void AnInvalidThemeColorIsIgnoredWithoutThrowing(int binaryTheme, double tint)
    {
        byte[] parameters = new byte[76];
        parameters[3] = 2;

        int cursor = 6;
        parameters[cursor] = 0x02;
        cursor += 11;
        parameters[cursor] = 0x03;
        cursor += 11;

        WriteThemeColor(parameters.AsSpan(cursor + 8), binaryTheme, tint);
        cursor += 24;
        WriteThemeColor(parameters.AsSpan(cursor + 8), 0, 0);

        ConditionalFormatRule rule = ReadRule(0x03, parameters);

        Assert.True(rule.ColorScale!.Colors[0].IsAuto);
        Assert.False(rule.ColorScale.Colors[1].IsAuto);
    }

    [Fact]
    public void AGradedRuleKeepsItsActivityFormulaAndNormalizesStopIfTrue()
    {
        byte[] parameters = new byte[28];
        parameters[3] = 0x02;
        parameters[4] = 10;
        parameters[5] = 90;
        BinaryPrimitives.WriteUInt32LittleEndian(parameters.AsSpan(6), 0x02);
        parameters[10] = 0x63;
        parameters[11] = 0x8E;
        parameters[12] = 0xC6;
        parameters[22] = 0x02;
        parameters[25] = 0x03;

        ConditionalFormatRule rule = ReadRule(
            0x04,
            parameters,
            activityFormula: Number(1),
            flags: 0x02);                   // Forbidden for graded rules; reader normalizes it.

        Assert.Equal(["1"], rule.Formulas);
        Assert.False(rule.StopIfTrue);
    }

    [Theory]
    [InlineData(0x04, ConditionalValueType.Percent)]
    [InlineData(0x05, ConditionalValueType.Percentile)]
    public void AFormulaValueObjectKeepsItsPercentSemantics(byte binaryType, ConditionalValueType expectedType)
    {
        byte[] tokens = Number(25);
        byte[] parameters = new byte[28 + tokens.Length];
        parameters[3] = 0x02;
        parameters[4] = 10;
        parameters[5] = 90;
        BinaryPrimitives.WriteUInt32LittleEndian(parameters.AsSpan(6), 0x02);
        parameters[10] = 0x63;
        parameters[11] = 0x8E;
        parameters[12] = 0xC6;

        int cursor = 22;
        parameters[cursor] = binaryType;
        BinaryPrimitives.WriteUInt16LittleEndian(parameters.AsSpan(cursor + 1), (ushort)tokens.Length);
        tokens.CopyTo(parameters, cursor + 3);
        cursor += 3 + tokens.Length;
        parameters[cursor] = 0x03;          // Maximum, with an empty formula.

        ConditionalFormatRule rule = ReadRule(0x04, parameters);

        Assert.Equal(new ConditionalValue(expectedType, "25"), rule.DataBar!.Minimum);
    }

    [Theory]
    [InlineData(0x00)]
    [InlineData(0x06)]
    [InlineData(0x07)]
    [InlineData(0xFF)]
    public void AnInvalidOrEmptyFormulaValueObjectIsReportedAndSkipped(byte binaryType)
    {
        byte[] parameters = new byte[28];
        parameters[3] = 0x02;
        parameters[4] = 10;
        parameters[5] = 90;
        parameters[22] = binaryType;         // Empty formula for every type, including 0x07.
        parameters[25] = 0x03;

        var globals = new XlsGlobals { Version = BiffVersion.Biff8 };
        var collector = new ConditionalFormatCollector();
        collector.Attach(globals);
        collector.OnCondFmt(Group(), isVersion12: true);
        collector.OnCf12(
            Rule(0x04, parameters, activityFormula: null, flags: 0),
            new PtgDecoder(globals, ["Sheet1"]),
            sheetIndex: 0);

        var warnings = new List<WorkbookWarning>();
        Worksheet sheet = Workbook.Create().Sheets.Add("Sheet1");
        collector.Finish(sheet, warnings.Add);

        Assert.Empty(sheet.ConditionalFormats);
        Assert.Contains(warnings, static warning => warning.Code == "CfRule12Skipped");
    }

    private static ConditionalFormatRule ReadRule(
        byte type, byte[] parameters, byte[]? activityFormula = null, byte flags = 0)
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff8 };
        var collector = new ConditionalFormatCollector();
        collector.Attach(globals);
        collector.OnCondFmt(Group(), isVersion12: true);
        collector.OnCf12(
            Rule(type, parameters, activityFormula, flags),
            new PtgDecoder(globals, ["Sheet1"]),
            sheetIndex: 0);

        Worksheet sheet = Workbook.Create().Sheets.Add("Sheet1");
        collector.Finish(sheet, _ => { });
        return Assert.Single(sheet.ConditionalFormats);
    }

    private static byte[] Group()
    {
        // Twelve-byte future-record header, the twelve-byte CondFmt header, then one A1 range.
        byte[] body = new byte[34];
        BinaryPrimitives.WriteUInt16LittleEndian(body.AsSpan(24), 1);
        return body;
    }

    private static byte[] Rule(byte type, byte[] parameters, byte[]? activityFormula, byte flags)
    {
        // The graded parameters begin after the future-record header, empty format/formulas,
        // the activity formula, flags, priority, template and the required 16-byte template
        // parameter block. With no activity formula, rgbCT starts at byte 48.
        int formulaLength = activityFormula?.Length ?? 0;
        int flagsOffset = 26 + formulaLength;
        int parametersOffset = flagsOffset + 22;
        byte[] body = new byte[parametersOffset + parameters.Length];
        body[12] = type;
        BinaryPrimitives.WriteUInt16LittleEndian(body.AsSpan(24), (ushort)formulaLength);
        activityFormula?.CopyTo(body, 26);
        body[flagsOffset] = flags;
        BinaryPrimitives.WriteUInt16LittleEndian(body.AsSpan(flagsOffset + 1), 1);
        body[flagsOffset + 5] = 16;          // cbTemplateParm.
        parameters.CopyTo(body, parametersOffset);
        return body;
    }

    private static void WriteThemeColor(Span<byte> body, int theme, double tint)
    {
        BinaryPrimitives.WriteUInt32LittleEndian(body, 0x03);       // XCLRTHEMED.
        BinaryPrimitives.WriteUInt32LittleEndian(body[4..], (uint)theme);
        BinaryPrimitives.WriteDoubleLittleEndian(body[8..], tint); // numTint.
    }

    private static byte[] Number(double value)
    {
        byte[] tokens = new byte[9];
        tokens[0] = 0x1F;                   // PtgNum.
        BinaryPrimitives.WriteDoubleLittleEndian(tokens.AsSpan(1), value);
        return tokens;
    }

    private static void WriteThreshold(byte[] body, ref int offset, double number, bool inclusive)
    {
        body[offset] = 0x04;                // CFVO percent, with an empty formula.
        BinaryPrimitives.WriteDoubleLittleEndian(body.AsSpan(offset + 3), number);
        offset += 11;
        body[offset] = inclusive ? (byte)1 : (byte)0;
        offset += 5;                        // fEqual and four unused bytes.
    }
}
