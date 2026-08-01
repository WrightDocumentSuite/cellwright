using Cellwright.Model;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;
using Cellwright.Xls.Xlsb;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The parts of the binary format that no file in the corpus happens to use: conditional
/// formatting, tables, and the structured references that name them.
/// </summary>
/// <remarks>
/// The fixtures here are built byte by byte rather than taken from a real workbook, and that
/// is a weaker kind of evidence: it shows the reader agrees with this file's reading of the
/// specification, not with Excel. Where the corpus does cover a feature — values, formulas,
/// comments, validation — the cross-format oracle is the check that counts, and these are only
/// the corners it cannot reach.
/// </remarks>
public sealed class XlsbFeatureTests
{
    /// <summary>
    /// A cell-value rule with two bounds. The binary package splits what the older format kept
    /// in one number into a type and a template, and the comparison arrives in a third field.
    /// </summary>
    [Fact]
    public void ACellValueRuleKeepsItsComparisonAndBounds()
    {
        Worksheet sheet = Read(new Biff12Builder()
            .Record(Biff12Records.BeginSheetData)
            .Record(Biff12Records.EndSheetData)
            .Record(Biff12Records.BeginConditionalFormatting, Group("B2:B10"))
            .Record(Biff12Records.BeginCFRule, CellIsBetween(low: 10, high: 20))
            .Record(Biff12Records.EndCFRule)
            .Record(Biff12Records.EndConditionalFormatting)
            .ToArray());

        ConditionalFormatRule rule = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.CellIs, rule.Type);
        Assert.Equal(ConditionalOperator.Between, rule.Operator);
        Assert.Equal(["10", "20"], rule.Formulas);
        Assert.Equal("B2:B10", Assert.Single(rule.Ranges).ToString());
    }

    /// <summary>
    /// A graded rule whose detail records never arrive has to be counted and reported rather
    /// than turned into a rule with no scale in it.
    /// </summary>
    [Fact]
    public void ARuleWithNoModelIsReportedRatherThanInvented()
    {
        var warnings = new List<WorkbookWarning>();
        Worksheet sheet = Read(
            new Biff12Builder()
                .Record(Biff12Records.BeginConditionalFormatting, Group("A1:A9"))
                .Record(Biff12Records.BeginCFRule, Graded(type: 3, template: 2))
                .Record(Biff12Records.EndCFRule)
                .Record(Biff12Records.EndConditionalFormatting)
                .ToArray(),
            warnings.Add);

        Assert.Empty(sheet.ConditionalFormats);
        Assert.Contains(warnings, static warning => warning.Code == "CfRule12Skipped");
    }

    /// <summary>A two-color scale: two value stops, two colors (MS-XLSB 2.4.28).</summary>
    [Fact]
    public void AColourScaleKeepsItsStopsAndColors()
    {
        Worksheet sheet = Read(new Biff12Builder()
            .Record(Biff12Records.BeginConditionalFormatting, Group("A1:A9"))
            .Record(Biff12Records.BeginCFRule, Graded(type: 3, template: 2))
            .Record(Biff12Records.BeginColorScale)
            .Record(Biff12Records.CFVO, Cfvo(type: 2))
            .Record(Biff12Records.CFVO, Cfvo(type: 3))
            .Record(Biff12Records.Color, Rgb(0xFF, 0x00, 0x00))
            .Record(Biff12Records.Color, Rgb(0x00, 0xB0, 0x50))
            .Record(Biff12Records.EndColorScale)
            .Record(Biff12Records.EndCFRule)
            .Record(Biff12Records.EndConditionalFormatting)
            .ToArray());

        ConditionalFormatRule rule = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.ColorScale, rule.Type);
        Assert.NotNull(rule.ColorScale);
        Assert.Equal([ConditionalValue.Minimum, ConditionalValue.Maximum], rule.ColorScale!.Values);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0x00, 0x00), rule.ColorScale.Colors[0]);
        Assert.Equal(ExcelColor.FromRgb(0x00, 0xB0, 0x50), rule.ColorScale.Colors[1]);
    }

    /// <summary>A data bar: length limits and the show flag, two stops, one color (2.4.43).</summary>
    [Fact]
    public void ADataBarKeepsItsStopsAndColor()
    {
        Worksheet sheet = Read(new Biff12Builder()
            .Record(Biff12Records.BeginConditionalFormatting, Group("B1:B9"))
            .Record(Biff12Records.BeginCFRule, Graded(type: 4, template: 3))
            .Record(Biff12Records.BeginDatabar, [0, 100, 0])
            .Record(Biff12Records.CFVO, Cfvo(type: 1, number: 5))
            .Record(Biff12Records.CFVO, Cfvo(type: 5, number: 90))
            .Record(Biff12Records.Color, Rgb(0x63, 0x8E, 0xC6))
            .Record(Biff12Records.EndDatabar)
            .Record(Biff12Records.EndCFRule)
            .Record(Biff12Records.EndConditionalFormatting)
            .ToArray());

        ConditionalFormatRule rule = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.DataBar, rule.Type);
        Assert.NotNull(rule.DataBar);
        Assert.Equal(ConditionalValue.Number(5), rule.DataBar!.Minimum);
        Assert.Equal(ConditionalValue.Percentile(90), rule.DataBar.Maximum);
        Assert.False(rule.DataBar.ShowValue);
        Assert.Equal(ExcelColor.FromRgb(0x63, 0x8E, 0xC6), rule.DataBar.Color);
    }

    /// <summary>An icon set: the set number, the display bits and a stop per icon (2.4.91).</summary>
    [Fact]
    public void AnIconSetKeepsItsNameAndThresholds()
    {
        var iconSet = new Biff12Builder.Body()
            .Int32(3)                       // iSet: KPI3TRAFFICLIGHTS1.
            .Byte(0x04)                     // fReverse set, fIcon clear.
            .Byte(0x00)
            .ToArray();

        Worksheet sheet = Read(new Biff12Builder()
            .Record(Biff12Records.BeginConditionalFormatting, Group("C1:C9"))
            .Record(Biff12Records.BeginCFRule, Graded(type: 6, template: 5))
            .Record(Biff12Records.BeginIconSet, iconSet)
            .Record(Biff12Records.CFVO, Cfvo(type: 4, number: 0))
            .Record(Biff12Records.CFVO, Cfvo(type: 4, number: 33))
            .Record(Biff12Records.CFVO, Cfvo(type: 4, number: 67))
            .Record(Biff12Records.EndIconSet)
            .Record(Biff12Records.EndCFRule)
            .Record(Biff12Records.EndConditionalFormatting)
            .ToArray());

        ConditionalFormatRule rule = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal(ConditionalFormatType.IconSet, rule.Type);
        Assert.NotNull(rule.IconSet);
        Assert.Equal("3TrafficLights1", rule.IconSet!.Name);
        Assert.Equal(
            [ConditionalValue.Percent(0), ConditionalValue.Percent(33), ConditionalValue.Percent(67)],
            rule.IconSet.Values);
        Assert.True(rule.IconSet.ShowValue);
        Assert.True(rule.IconSet.Reverse);
    }

    /// <summary>A table gives up its name, its range and the names of its columns.</summary>
    [Fact]
    public void ATableKeepsItsNameAndColumns()
    {
        Worksheet sheet = NewSheet();
        var byId = new Dictionary<uint, ExcelTable>();
        XlsbTablePart.Read(TablePart(), sheet, byId);

        ExcelTable table = Assert.Single(sheet.Tables);
        Assert.Equal("Sales", table.Name);
        Assert.Equal("A1:C5", table.Range.ToString());
        Assert.Equal(["Region", "Quarter", "Amount"], table.Columns);
        Assert.True(table.HasHeaderRow);
    }

    /// <summary>
    /// A structured reference is one token here, not the text a person typed. Without the table
    /// part it cannot be rendered at all, which is why reading tables is what turns the formula
    /// from a warning into an expression.
    /// </summary>
    [Theory]
    [InlineData(0x0001, 0, 0, 0, "Sales[Region]")]
    [InlineData(0x0001, 2, 2, 0, "Sales[Amount]")]
    [InlineData(0x0002, 0, 2, 0, "Sales[[Region]:[Amount]]")]
    [InlineData(0x0000, 0, 0, 0x01, "Sales[#All]")]
    [InlineData(0x0001, 2, 2, 0x02, "Sales[[#Headers],[Amount]]")]
    [InlineData(0x0001, 2, 2, 0x10, "Sales[[#This Row],[Amount]]")]
    public void AStructuredReferenceIsRenderedFromTheTable(int columns, int first, int last, int rowType, string expected)
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        var decoder = new PtgDecoder(globals, ["Sheet1"]);
        Worksheet sheet = NewSheet();
        XlsbTablePart.Read(TablePart(), sheet, decoder.Tables);

        byte[] tokens = PtgList(columns, first, last, rowType);

        Assert.Equal(expected, decoder.Decode(tokens, new CellRef(2, 5), sheetIndex: 0));
    }

    /// <summary>
    /// A reference into a table the file never defined has no text, and saying so is the only
    /// honest answer: rendering the number would produce a formula that means nothing.
    /// </summary>
    [Fact]
    public void AReferenceToAnUnknownTableDeclines()
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        var decoder = new PtgDecoder(globals, ["Sheet1"]);

        Assert.Null(decoder.Decode(PtgList(0x0001, 0, 0, 0), new CellRef(2, 5), sheetIndex: 0));
    }

    private static Worksheet NewSheet() => Workbook.Create().Sheets.Add("Sheet1");

    private static Worksheet Read(byte[] part, Action<WorkbookWarning>? warn = null)
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        var decoder = new PtgDecoder(globals, ["Sheet1"]);
        Worksheet sheet = NewSheet();
        Action<WorkbookWarning> report = warn ?? (_ => { });

        var sink = new WorksheetContentSink(sheet, styleMap: [], relationships: [], preserveMarkup: false, report);
        new XlsbSheetParser(globals, decoder, sheetIndex: 0, report).Parse(part, sink);
        return sheet;
    }

    /// <summary>The header of a conditional formatting group: a rule count, a flag, the ranges.</summary>
    private static byte[] Group(string range)
    {
        RangeRef parsed = RangeRef.Parse(range);
        return new Biff12Builder.Body()
            .UInt32(1)
            .UInt32(0)
            .Int32(1)
            .Range(parsed.Start.Row - 1, parsed.End.Row - 1, parsed.Start.Column - 1, parsed.End.Column - 1)
            .ToArray();
    }

    /// <summary>A CF_TYPE_CELLIS rule comparing against two numbers (MS-XLSB 2.4.23).</summary>
    private static byte[] CellIsBetween(int low, int high)
    {
        byte[] first = Number(low);
        byte[] second = Number(high);
        var body = new Biff12Builder.Body()
            .UInt32(1)                      // iType: CF_TYPE_CELLIS.
            .UInt32(0)                      // iTemplate: CF_TEMPLATE_EXPR.
            .Int32(-1)                      // dxfId: no differential format.
            .Int32(1)                       // iPri.
            .UInt32(1)                      // iParam: CF_OPER_BN, between.
            .UInt32(0)                      // reserved1.
            .UInt32(0)                      // reserved2.
            .UInt32(0)                      // flags.
            .UInt32((uint)first.Length)     // cbFmla1: nonzero means the formula is present.
            .UInt32((uint)second.Length)
            .UInt32(0)                      // cbFmla3.
            .NoText();                      // strParam.

        body.UInt32((uint)first.Length);
        foreach (byte value in first)
            body.Byte(value);
        body.UInt32((uint)second.Length);
        foreach (byte value in second)
            body.Byte(value);

        return body.ToArray();
    }

    /// <summary>A graded rule header — color scale, data bar or icon set (MS-XLSB 2.4.23).</summary>
    private static byte[] Graded(int type, int template) => new Biff12Builder.Body()
        .UInt32((uint)type)                 // iType.
        .UInt32((uint)template)             // iTemplate.
        .Int32(-1)                          // dxfId: graded rules have none.
        .Int32(1)                           // iPri.
        .UInt32(0)
        .UInt32(0)
        .UInt32(0)
        .UInt32(0)
        .UInt32(0)
        .UInt32(0)
        .UInt32(0)
        .NoText()
        .ToArray();

    /// <summary>A <c>BrtCFVO</c> (MS-XLSB 2.4.334): the type, the number, the flags, no formula.</summary>
    private static byte[] Cfvo(int type, double number = 0)
    {
        var body = new Biff12Builder.Body().Int32(type);
        foreach (byte piece in BitConverter.GetBytes(number))
            body.Byte(piece);
        return body
            .UInt32(type >= 4 ? 1u : 0u)    // fSaveGTE, set the way icon sets require.
            .UInt32(1)                      // fGTE.
            .UInt32(0)                      // cbFmla.
            .ToArray();
    }

    /// <summary>A <c>BrtColor</c> (MS-XLSB 2.4.337) stating an explicit RGB.</summary>
    private static byte[] Rgb(byte red, byte green, byte blue) => new Biff12Builder.Body()
        .Byte(0x02)                         // xColorType: RGB.
        .Byte(0)                            // index.
        .Byte(0).Byte(0)                    // nTintAndShade.
        .Byte(red).Byte(green).Byte(blue).Byte(0xFF)
        .ToArray();

    /// <summary>A <c>PtgNum</c>: the token byte and eight bytes of double.</summary>
    private static byte[] Number(double value)
    {
        var body = new Biff12Builder.Body().Byte(0x1F);
        foreach (byte piece in BitConverter.GetBytes(value))
            body.Byte(piece);
        return body.ToArray();
    }

    /// <summary>A table part holding one three-column table (MS-XLSB 2.4.100 and 2.4.101).</summary>
    private static byte[] TablePart()
    {
        var list = new Biff12Builder.Body()
            .Range(0, 4, 0, 2)              // rfxList: A1:C5.
            .UInt32(0)                      // lt: LTRANGE.
            .UInt32(7)                      // idList.
            .UInt32(1)                      // crwHeader.
            .UInt32(0)                      // crwTotals.
            .UInt32(0)                      // flags.
            .Pad(24)                        // Six differential format identifiers.
            .UInt32(0)                      // dwConnID.
            .NoText()                       // stName.
            .Text("Sales")                  // stDisplayName, which formulas use.
            .NoText()                       // stComment.
            .NoText()                       // stStyleHeader.
            .NoText()                       // stStyleData.
            .NoText();                      // stStyleAgg.

        var builder = new Biff12Builder().Record(Biff12Records.BeginList, list.ToArray());
        builder.Record(Biff12Records.BeginListCols);
        foreach (string name in new[] { "Region", "Quarter", "Amount" })
            builder.Record(Biff12Records.BeginListCol, Column(name)).Record(Biff12Records.EndListCol);

        return builder
            .Record(Biff12Records.EndListCols)
            .Record(Biff12Records.EndList)
            .ToArray();
    }

    private static byte[] Column(string caption) => new Biff12Builder.Body()
        .UInt32(1)                          // idField.
        .UInt32(0)                          // ilta.
        .UInt32(0xFFFFFFFF)                 // nDxfHdr.
        .UInt32(0xFFFFFFFF)                 // nDxfInsertRow.
        .UInt32(0xFFFFFFFF)                 // nDxfAgg.
        .UInt32(0)                          // idqsif.
        .NoText()                           // stName.
        .Text(caption)                      // stCaption, which the header cell shows.
        .NoText()                           // stTotal.
        .NoText()                           // stStyleHeader.
        .ToArray();

    /// <summary>A <c>PtgList</c> token (MS-XLSB 2.5.98.52): fourteen bytes in all.</summary>
    private static byte[] PtgList(int columns, int first, int last, int rowType)
    {
        int flags = (columns & 0x03) | ((rowType & 0x1F) << 2);
        return new Biff12Builder.Body()
            .Byte(0x18)                     // ptg.
            .Byte(0x19)                     // eptg.
            .UInt16(0)                      // ixti.
            .UInt16(flags)
            .UInt32(7)                      // listIndex, matching idList above.
            .UInt16(first)
            .UInt16(last)
            .ToArray();
    }
}

/// <summary>Record numbers the fixtures above write, from the enumeration in MS-XLSB 2.3.</summary>
internal static class Biff12Records
{
    public const int BeginSheetData = 145;
    public const int EndSheetData = 146;
    public const int BeginConditionalFormatting = 461;
    public const int EndConditionalFormatting = 462;
    public const int BeginCFRule = 463;
    public const int EndCFRule = 464;
    public const int BeginIconSet = 465;
    public const int EndIconSet = 466;
    public const int BeginDatabar = 467;
    public const int EndDatabar = 468;
    public const int BeginColorScale = 469;
    public const int EndColorScale = 470;
    public const int CFVO = 471;
    public const int Color = 564;
    public const int BeginList = 343;
    public const int EndList = 344;
    public const int BeginListCols = 345;
    public const int EndListCols = 346;
    public const int BeginListCol = 347;
    public const int EndListCol = 348;
}
