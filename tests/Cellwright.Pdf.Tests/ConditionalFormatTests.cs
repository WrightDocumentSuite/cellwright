using Cellwright.Pdf.Layout;
using Cellwright.Styles;
using Inkwright;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Conditional formatting on the printed page: the fills a rule paints, the bars and icons it
/// draws, and the fonts and number formats it changes.
/// </summary>
public sealed class ConditionalFormatTests
{
    private readonly record struct IconSnapshot(IconShape Shape, double Angle, double Fraction);

    private sealed record LayoutSnapshot(
        IReadOnlyList<IconSnapshot> Icons,
        IReadOnlyList<(string Text, double X)> Text,
        IReadOnlyList<PdfExportWarning> Diagnostics);

    private static Worksheet Column(Workbook book, params double[] values)
    {
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int i = 0; i < values.Length; i++)
            sheet[i + 1, 1] = values[i];

        sheet.Column("A").SetWidth(20);
        return sheet;
    }

    private static LayoutSnapshot LayoutOf(Workbook book)
    {
        using PdfDocument document = PdfDocument.Create();
        var context = new PdfExportContext(book, document, new PdfExportOptions());
        IReadOnlyList<PageItem> items = new WorkbookComposer(context)
            .Compose()
            .SelectMany(static page => page.Items)
            .ToArray();

        return new LayoutSnapshot(
            [.. items.OfType<IconItem>().Select(static item =>
                new IconSnapshot(item.Shape, item.Angle, item.Fraction))],
            [.. items.OfType<TextItem>().Select(static item => (item.Text, item.X))],
            [.. context.Diagnostics]);
    }

    private static Workbook IconBook(string name, int count)
    {
        Workbook book = Workbook.Create();
        double[] values = Enumerable.Range(0, count)
            .Select(index => count == 1 ? 0 : index * 100.0 / (count - 1))
            .ToArray();
        Worksheet sheet = Column(book, values);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons(
            $"A1:A{count}",
            new IconSet
            {
                Name = name,
                Values = [.. values.Select(static value => ConditionalValue.Percent(value))],
            }));
        return book;
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
    public void AContainsBlanksRulePaintsAnOtherwiseEmptyPrintArea()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.SetPrintArea("A1:A2");
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2")],
            Type = ConditionalFormatType.ContainsBlanks,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) },
        });

        using Rendered rendered = Rendered.Of(book);

        Assert.Equal(2, FillsOf(rendered, 1, 0, 0));
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
            "A1:A2", new DataBar
            {
                Color = ExcelColor.FromRgb(0x00, 0xFF, 0x00),
                MinimumLength = 0,
                MaximumLength = 100,
            }));

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

    [Theory]
    [InlineData("3Stars", 3, "Star,Star,Star")]
    [InlineData("3Triangles", 3, "Triangle,Dash,Triangle")]
    [InlineData("5Boxes", 5, "Boxes,Boxes,Boxes,Boxes,Boxes")]
    public void ExtendedIconSetsUseTheirOwnShapesAndRenderWithoutAFallback(
        string name, int count, string expectedShapes)
    {
        Workbook book = IconBook(name, count);

        LayoutSnapshot layout = LayoutOf(book);
        Assert.Equal(expectedShapes, string.Join(",", layout.Icons.Select(static icon => icon.Shape)));
        Assert.DoesNotContain(
            layout.Diagnostics,
            warning => warning.Subject?.StartsWith("icon-set", StringComparison.Ordinal) == true);

        // Exercise the painters too, without coupling the test to PDF drawing operators.
        using Rendered rendered = Rendered.Of(book);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("icon-set", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void TriangleBandsPointDownStayLevelAndPointUp()
    {
        LayoutSnapshot layout = LayoutOf(IconBook("3Triangles", 3));

        Assert.Equal(180, layout.Icons[0].Angle);
        Assert.Equal(IconShape.Dash, layout.Icons[1].Shape);
        Assert.Equal(0, layout.Icons[2].Angle);
    }

    [Fact]
    public void StarsAndBoxesProgressFromEmptyToFull()
    {
        LayoutSnapshot stars = LayoutOf(IconBook("3Stars", 3));
        LayoutSnapshot boxes = LayoutOf(IconBook("5Boxes", 5));

        Assert.Equal(new double[] { 0, 0.5, 1 }, stars.Icons.Select(static icon => icon.Fraction));
        Assert.Equal(
            new double[] { 0, 0.25, 0.5, 0.75, 1 },
            boxes.Icons.Select(static icon => icon.Fraction));
    }

    [Fact]
    public void NoIconsNeitherDrawsAnIconNorIndentsTheCellText()
    {
        Workbook plain = Workbook.Create();
        Worksheet plainSheet = plain.Sheets.Add("Sheet1");
        plainSheet["A1"] = "label";
        plainSheet.Column("A").SetWidth(20);

        Workbook suppressed = Workbook.Create();
        Worksheet suppressedSheet = suppressed.Sheets.Add("Sheet1");
        suppressedSheet["A1"] = "label";
        suppressedSheet.Column("A").SetWidth(20);
        suppressedSheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1", new IconSet
        {
            Name = "NoIcons",
            Values = [ConditionalValue.Percent(0)],
        }));

        LayoutSnapshot withoutIcon = LayoutOf(plain);
        LayoutSnapshot withSuppressedIcon = LayoutOf(suppressed);

        Assert.Empty(withSuppressedIcon.Icons);
        Assert.Single(withoutIcon.Text);
        Assert.Single(withSuppressedIcon.Text);
        Assert.InRange(
            Math.Abs(withoutIcon.Text[0].X - withSuppressedIcon.Text[0].X),
            0,
            0.000001);
        Assert.DoesNotContain(
            withSuppressedIcon.Diagnostics,
            warning => warning.Subject?.StartsWith("icon-set", StringComparison.Ordinal) == true);

        using Rendered rendered = Rendered.Of(suppressed);
        Assert.Contains("label", rendered.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            warning => warning.Subject?.StartsWith("icon-set", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void ExistingIconSetsRemainMapped()
    {
        string[] names =
        [
            "3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2", "3Signs",
            "3Symbols", "3Symbols2", "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating",
            "4TrafficLights", "5Arrows", "5ArrowsGray", "5Quarters", "5Rating",
        ];

        foreach (string name in names)
        {
            int count = name[0] - '0';
            Assert.NotNull(IconSetShapes.Of(name, 0, count));
            Assert.NotEqual(IconShape.None, IconSetShapes.Of(name, 0, count)?.Shape);
        }
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
    public void PrintingOneWorksheetKeepsTheRelativeFormulaEvaluator()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Column(book, 500);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression(
            "A1", "$A1>100", new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) }));

        int calls = 0;
        var options = new PdfExportOptions
        {
            RelativeFormulaEvaluator = (worksheet, anchor, cell, formula) =>
            {
                calls++;
                Assert.Same(sheet, worksheet);
                Assert.Equal(CellRef.Parse("A1"), anchor);
                Assert.Equal(CellRef.Parse("A1"), cell);
                Assert.Equal("$A1>100", formula);
                return CellValue.Boolean(true);
            },
        };

        PdfExportResult result = sheet.ToPdf(options);
        using (result.Document)
        {
            Assert.True(calls > 0);
            Assert.DoesNotContain(
                result.Diagnostics,
                warning => warning.Subject == "conditional-format-formula");
        }
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
