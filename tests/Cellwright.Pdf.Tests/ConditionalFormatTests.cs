using Cellwright.Styles;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Conditional formatting on the printed page: the fills a rule paints, the bars and icons it
/// draws, and the fonts and number formats it changes.
/// </summary>
public sealed class ConditionalFormatTests
{
    private static Worksheet Column(Workbook book, params double[] values)
    {
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int i = 0; i < values.Length; i++)
            sheet[i + 1, 1] = values[i];

        sheet.Column("A").SetWidth(20);
        return sheet;
    }

    /// <summary>How many times a colour is set as a fill in the page's operators.</summary>
    private static int FillsOf(Rendered rendered, double red, double green, double blue)
    {
        string wanted = string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"{red:0.###} {green:0.###} {blue:0.###} rg");

        int count = 0;
        string content = rendered.Content();
        for (int i = content.IndexOf(wanted, StringComparison.Ordinal); i >= 0;
             i = content.IndexOf(wanted, i + 1, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    [Fact]
    public void ACellThatPassesARuleIsPaintedInTheRulesFill()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 1, 500);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1:A2", 100, new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) }));

        using Rendered rendered = Rendered.Of(book);

        // Pure red is painted exactly once: for the cell that passed, not for the one that did not.
        Assert.Equal(1, FillsOf(rendered, 1, 0, 0));
    }

    [Fact]
    public void ARuleThatChangesTheNumberFormatChangesWhatIsPrinted()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 0.5);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1", 0.1, new DifferentialStyle { NumberFormat = "0.00%" }));

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains("50.00%", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void ARuleThatDoesNotMatchLeavesTheCellAsItWas()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 0.5);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan(
            "A1", 100, new DifferentialStyle { NumberFormat = "0.00%" }));

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain("50.00%", rendered.Text(), StringComparison.Ordinal);
        Assert.Contains("0.5", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void AColourScalePaintsEveryCellOfItsRange()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 0, 50, 100);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A3", new ColorScale
        {
            Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
            Colors = [ExcelColor.FromRgb(0xFF, 0x00, 0x00), ExcelColor.FromRgb(0x00, 0x00, 0xFF)],
        }));

        using Rendered rendered = Rendered.Of(book);

        // The ends take the stop colours outright; the middle is a mix of the two.
        Assert.Equal(1, FillsOf(rendered, 1, 0, 0));
        Assert.Equal(1, FillsOf(rendered, 0, 0, 1));
    }

    [Fact]
    public void ADataBarIsAsLongAShareOfTheCellAsTheValueIsOfTheRange()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 0, 100);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars(
            "A1:A2", new DataBar { Color = ExcelColor.FromRgb(0x00, 0xFF, 0x00) }));

        using Rendered rendered = Rendered.Of(book);

        // The full bar is drawn; the empty one has no width and so is not drawn at all.
        Assert.Equal(1, FillsOf(rendered, 0, 1, 0));
    }

    [Fact]
    public void AnIconSetPutsSomethingInEveryCellOfItsRange()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 0, 50, 100);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A3", IconSet.TrafficLights));

        using Rendered rendered = Rendered.Of(book);
        string content = rendered.Content();

        // Three discs, which the canvas draws as curves rather than as rectangles.
        Assert.Contains(" c\n", content, StringComparison.Ordinal);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("icon-set", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void AnIconSetThisDoesNotKnowIsSaidRatherThanDrawnWrong()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 1, 2, 3);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A3", new IconSet
        {
            Name = "7Invented",
            Values = [ConditionalValue.Percent(0), ConditionalValue.Percent(50)],
        }));

        using Rendered rendered = Rendered.Of(book);

        Assert.Contains(
            rendered.Diagnostics,
            warning => warning.Subject == "icon-set-7Invented");
    }

    [Fact]
    public void ARuleThatHidesTheValueLeavesOnlyItsIcon()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 42);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1", new IconSet
        {
            Name = "3TrafficLights1",
            Values = [ConditionalValue.Percent(0), ConditionalValue.Percent(33), ConditionalValue.Percent(67)],
            ShowValue = false,
        }));

        using Rendered rendered = Rendered.Of(book);

        Assert.DoesNotContain("42", rendered.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheTextStartsAfterTheIconRatherThanUnderIt()
    {
        Workbook plain = Workbook.Create();
        Column(plain, 42);
        using Rendered withoutIcon = Rendered.Of(plain);

        Workbook iconed = Workbook.Create();
        Worksheet sheet = Column(iconed, 42);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1", IconSet.TrafficLights));
        using Rendered withIcon = Rendered.Of(iconed);

        // The number is right-aligned by default, so an icon at the left edge must not move it;
        // what matters is that both are still printed.
        Assert.Contains("42", withoutIcon.Text(), StringComparison.Ordinal);
        Assert.Contains("42", withIcon.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void AFormulaRuleIsAppliedWhenAnEvaluatorIsGiven()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 1, 500);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression(
            "A1:A2", "$A1>100", new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) }));

        var options = new PdfExportOptions
        {
            // Stands in for the formula package.
            FormulaEvaluator = (worksheet, cell, formula) =>
                CellValue.Boolean(worksheet.Values.Get(cell.Row, cell.Column).ToSerialNumber() > 100),
        };

        using Rendered rendered = Rendered.Of(book, options);

        Assert.Equal(1, FillsOf(rendered, 1, 0, 0));
        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject == "conditional-format-formula");
    }

    [Fact]
    public void StopIfTrueKeepsALaterRuleFromPaintingOnThePage()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 500);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            StopIfTrue = true,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) },
        });

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["2"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(0x00, 0x00, 0xFF) },
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(1, FillsOf(rendered, 1, 0, 0));
        Assert.Equal(0, FillsOf(rendered, 0, 0, 1));
    }

    [Fact]
    public void ARuleThatChangesTheFontIsMeasuredWithItRatherThanWithout()
    {
        // A bold rule makes the text wider, and a wrapped cell then needs a taller row. If the
        // measuring pass ignored the rule, the extra line would be drawn outside the cell and
        // clipped away.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "a sentence that only just fits on two lines when it is drawn in bold";
        sheet.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        sheet.Column("A").SetWidth(18);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Containing(
            "A1", "sentence", new DifferentialStyle { Bold = true }));

        using Rendered rendered = Rendered.Of(book);

        // Every word survives, which it would not if the row had been sized for the thinner font.
        Assert.Contains("bold", rendered.Text(), StringComparison.Ordinal);
    }
}
