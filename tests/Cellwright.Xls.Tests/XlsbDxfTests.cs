using Cellwright.Model;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;
using Cellwright.Xls.Xlsb;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>Binary differential formats and their XLSB-to-XLSX conversion.</summary>
public sealed class XlsbDxfTests
{
    [Fact]
    public void ADirectNumberFormatWinsEvenWhenItsIdPropertyComesLater()
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        Workbook workbook = Workbook.Create();

        DifferentialStyle style = Assert.Single(
            XlsbStyles.Read(StylesWithDifferentialFormat(), globals, workbook.StylePool));

        Assert.Equal("0.0000", style.NumberFormat);
    }

    [Fact]
    public async Task AnOrdinaryRuleKeepsItsDxfWhenConvertedToXlsx()
    {
        Worksheet binary = Read(
            new Biff12Builder()
                .Record(Records.BeginConditionalFormatting, Group("B2:B10"))
                .Record(Records.BeginCFRule, CellIsBetween(low: 10, high: 20, styleId: 0))
                .Record(Records.EndCFRule)
                .Record(Records.EndConditionalFormatting)
                .ToArray(),
            StylesWithDifferentialFormat());

        ConditionalFormatRule sourceRule = Assert.Single(binary.ConditionalFormats);
        Assert.Equal(0, sourceRule.SourceStyleId);
        AssertDifferentialFormat(sourceRule.Style);

        using var buffer = new MemoryStream();
        await binary.Workbook.SaveAsync(
            buffer,
            new SaveOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        buffer.Position = 0;
        Workbook modern = await Workbook.LoadAsync(
            buffer,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        ConditionalFormatRule convertedRule = Assert.Single(modern.Sheets[0].ConditionalFormats);
        Assert.Equal(ConditionalFormatType.CellIs, convertedRule.Type);
        Assert.Equal(["10", "20"], convertedRule.Formulas);
        AssertDifferentialFormat(convertedRule.Style);
    }

    private static Worksheet Read(byte[] sheetPart, byte[] stylesPart)
    {
        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Sheet1");
        IReadOnlyList<DifferentialStyle> differentialStyles =
            XlsbStyles.Read(stylesPart, globals, workbook.StylePool);
        workbook.LoadedDifferentialStyles = differentialStyles;

        var decoder = new PtgDecoder(globals, [sheet.Name]);
        var sink = new WorksheetContentSink(
            sheet, styleMap: [], relationships: [], preserveMarkup: false, _ => { });
        new XlsbSheetParser(globals, decoder, sheetIndex: 0, _ => { }).Parse(sheetPart, sink);
        XlsbWorkbookReader.AttachConditionalStyles(sheet, differentialStyles);
        return sheet;
    }

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

    private static byte[] CellIsBetween(int low, int high, int styleId)
    {
        byte[] first = Number(low);
        byte[] second = Number(high);
        var body = new Biff12Builder.Body()
            .UInt32(1)                      // iType: CF_TYPE_CELLIS.
            .UInt32(0)                      // iTemplate: CF_TEMPLATE_EXPR.
            .Int32(styleId)                 // dxfId.
            .Int32(1)                       // iPri.
            .UInt32(1)                      // iParam: CF_OPER_BN, between.
            .UInt32(0)
            .UInt32(0)
            .UInt32(0)
            .UInt32((uint)first.Length)
            .UInt32((uint)second.Length)
            .UInt32(0)
            .NoText();

        ParsedFormula(body, first);
        ParsedFormula(body, second);
        return body.ToArray();
    }

    private static void ParsedFormula(Biff12Builder.Body body, byte[] tokens)
    {
        body.UInt32((uint)tokens.Length);
        foreach (byte value in tokens)
            body.Byte(value);
        body.UInt32(0);
    }

    private static byte[] Number(double value)
    {
        var body = new Biff12Builder.Body().Byte(0x1F);
        foreach (byte piece in BitConverter.GetBytes(value))
            body.Byte(piece);
        return body.ToArray();
    }

    /// <summary>
    /// One BrtDXF composed of independently-sized XFProp entries. The unmodelled font-size
    /// entry proves that its <c>cb</c> keeps the following number-format property aligned.
    /// </summary>
    private static byte[] StylesWithDifferentialFormat()
    {
        var dxf = new Biff12Builder.Body()
            .UInt16(0x8000)                 // fNewBorder.
            .UInt16(0)                      // XFProps.reserved.
            .UInt16(10);                    // XFProps.cprops.

        XfProperty(dxf, 0x0002, DxfColor(0xFF, 0xC7, 0xCE));
        XfProperty(dxf, 0x0005, DxfColor(0x9C, 0x00, 0x06));
        XfProperty(dxf, 0x0019, new Biff12Builder.Body().UInt16(0x02BC).ToArray());
        XfProperty(dxf, 0x001A, new Biff12Builder.Body().UInt16(1).ToArray());
        XfProperty(dxf, 0x001C, [1]);
        XfProperty(dxf, 0x001D, [1]);

        var border = new Biff12Builder.Body();
        foreach (byte value in DxfColor(0x11, 0x22, 0x33))
            border.Byte(value);
        XfProperty(dxf, 0x0006, border.UInt16(1).ToArray()); // thin top border.

        XfProperty(dxf, 0x0024, new Biff12Builder.Body().UInt32(240).ToArray());
        // The direct code is authoritative even when a producer places its id afterwards.
        XfProperty(dxf, 0x0026, LpWideString("0.0000"));
        XfProperty(dxf, 0x0029, new Biff12Builder.Body().UInt16(164).ToArray());

        return new Biff12Builder()
            .Record(Records.Fmt, new Biff12Builder.Body().UInt16(164).Text("0.00%").ToArray())
            .Record(Records.BeginDXFs, new Biff12Builder.Body().UInt32(1).ToArray())
            .Record(Records.Dxf, dxf.ToArray())
            .Record(Records.EndDXFs)
            .ToArray();
    }

    private static void XfProperty(Biff12Builder.Body target, int type, byte[] data)
    {
        target.UInt16(type).UInt16(data.Length + 4);
        foreach (byte value in data)
            target.Byte(value);
    }

    private static byte[] LpWideString(string value)
    {
        var body = new Biff12Builder.Body().UInt16(value.Length);
        foreach (char character in value)
            body.UInt16(character);
        return body.ToArray();
    }

    /// <summary>An XFPropColor: fValidRGBA, RGB kind, tint, then LongRGBA.</summary>
    private static byte[] DxfColor(byte red, byte green, byte blue) => new Biff12Builder.Body()
        .Byte((2 << 1) | 1)
        .Byte(0xFF)
        .UInt16(0)
        .Byte(red).Byte(green).Byte(blue).Byte(0xFF)
        .ToArray();

    private static void AssertDifferentialFormat(DifferentialStyle? style)
    {
        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.True(style.Underline);
        Assert.True(style.Strikethrough);
        Assert.Equal(ExcelColor.FromRgb(0x9C, 0x00, 0x06), style.FontColor);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0xC7, 0xCE), style.FillColor);
        Assert.Equal("0.0000", style.NumberFormat);
        Assert.NotNull(style.Border);
        Assert.Equal(BorderLineStyle.Thin, style.Border.Top.Line);
        Assert.Equal(ExcelColor.FromRgb(0x11, 0x22, 0x33), style.Border.Top.Color);
    }

    private static class Records
    {
        public const int Fmt = 44;
        public const int BeginConditionalFormatting = 461;
        public const int EndConditionalFormatting = 462;
        public const int BeginCFRule = 463;
        public const int EndCFRule = 464;
        public const int BeginDXFs = 505;
        public const int EndDXFs = 506;
        public const int Dxf = 507;
    }
}
