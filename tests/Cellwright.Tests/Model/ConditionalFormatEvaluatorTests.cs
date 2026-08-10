using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// What a sheet's conditional formatting rules actually come to for a cell: which ones match,
/// in what order they paint over each other, and what the graded ones work out to.
/// </summary>
public sealed class ConditionalFormatEvaluatorTests
{
    private static readonly DateOnly Today = new(2026, 7, 30);

    private static Worksheet Column(params double[] values)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int i = 0; i < values.Length; i++)
            sheet[i + 1, 1] = values[i];

        return sheet;
    }

    private static ConditionalFormatEvaluator Evaluator(
        Worksheet sheet, Func<string, CellRef, CellValue>? evaluate = null) =>
        new(sheet, evaluate, Today);

    private static CellRef At(int row) => new(row, 1);

    [Fact]
    public void AComparisonAgainstALiteralPaintsTheCellsThatPassIt()
    {
        Worksheet sheet = Column(5, 50, 500);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A3", 100, DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Equal(DifferentialStyle.Bad, evaluator.Of(At(3)).Style);
    }

    [Fact]
    public void BetweenIsInclusiveAtBothEnds()
    {
        Worksheet sheet = Column(9, 10, 15, 20, 21);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellBetween("A1:A5", 10, 20, DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.NotNull(evaluator.Of(At(2)).Style);
        Assert.NotNull(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
        Assert.Null(evaluator.Of(At(5)).Style);
    }

    [Fact]
    public void ALowerPriorityRuleSuppliesOnlyPropertiesTheHigherOneDoesNot()
    {
        Worksheet sheet = Column(500);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            Style = new DifferentialStyle { Bold = true, FontColor = ExcelColor.FromRgb(1, 2, 3) },
        });

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["2"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(9, 9, 9) },
        });

        DifferentialStyle? style = Evaluator(sheet).Of(At(1)).Style;

        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.Equal(ExcelColor.FromRgb(1, 2, 3), style.FontColor);
        Assert.Equal(ExcelColor.FromRgb(9, 9, 9), style.FillColor);
    }

    [Fact]
    public void BorderEdgesFromMatchingRulesMergeIndependently()
    {
        Worksheet sheet = Column(500);
        var higherLeft = new BorderEdge(BorderLineStyle.Thin, ExcelColor.FromRgb(1, 2, 3));
        var lowerLeft = new BorderEdge(BorderLineStyle.Dashed, ExcelColor.FromRgb(4, 5, 6));
        var lowerRight = new BorderEdge(BorderLineStyle.Double, ExcelColor.FromRgb(7, 8, 9));

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            Style = new DifferentialStyle { Border = new BorderStyle { Left = higherLeft } },
        });
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 2,
            Style = new DifferentialStyle
            {
                Border = new BorderStyle { Left = lowerLeft, Right = lowerRight },
            },
        });

        BorderStyle border = Assert.IsType<BorderStyle>(Evaluator(sheet).Of(At(1)).Style!.Border);

        Assert.Equal(higherLeft, border.Left);
        Assert.Equal(lowerRight, border.Right);
    }

    [Fact]
    public void AHigherPriorityRuleWinsWhenBothSetTheSameProperty()
    {
        Worksheet sheet = Column(500);
        ExcelColor higher = ExcelColor.FromRgb(1, 2, 3);
        ExcelColor lower = ExcelColor.FromRgb(9, 8, 7);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            Style = new DifferentialStyle { FillColor = higher },
        });
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = lower },
        });

        Assert.Equal(higher, Evaluator(sheet).Of(At(1)).Style!.FillColor);
    }

    [Fact]
    public void EqualPrioritiesKeepTheRulesCollectionOrder()
    {
        Worksheet sheet = Column(500);
        ExcelColor first = ExcelColor.FromRgb(1, 2, 3);

        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 7,
            Style = new DifferentialStyle { FillColor = first },
        });
        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 7,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(9, 8, 7) },
        });

        Assert.Equal(first, Evaluator(sheet).Of(At(1)).Style!.FillColor);
    }

    [Fact]
    public void AHigherPriorityDifferentialFillWinsOverALowerColourScale()
    {
        Worksheet sheet = Column(50);
        ExcelColor higher = ExcelColor.FromRgb(1, 2, 3);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            Style = new DifferentialStyle { FillColor = higher },
        });
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.ColorScale,
            Priority = 2,
            ColorScale = new ColorScale
            {
                Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
                Colors = [ExcelColor.FromRgb(9, 0, 0), ExcelColor.FromRgb(0, 0, 9)],
            },
        });

        ConditionalAppearance appearance = Evaluator(sheet).Of(At(1));

        Assert.Equal(higher, appearance.Style!.FillColor);
        Assert.True(appearance.ScaleColor.IsAuto);
    }

    [Fact]
    public void StopIfTrueKeepsLaterRulesFromPainting()
    {
        Worksheet sheet = Column(500);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 1,
            StopIfTrue = true,
            Style = new DifferentialStyle { Bold = true },
        });

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["2"],
            Priority = 2,
            Style = new DifferentialStyle { FillColor = ExcelColor.FromRgb(9, 9, 9) },
        });

        DifferentialStyle? style = Evaluator(sheet).Of(At(1)).Style;

        Assert.NotNull(style);
        Assert.True(style.Bold);
        Assert.True(style.FillColor.IsAuto);
    }

    [Fact]
    public void TextRulesLookAtWhatTheCellSays()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Overdue invoice";
        sheet["A2"] = "Paid";

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Containing("A1:A2", "overdue", DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
    }

    [Fact]
    public void BlankRulesUseExcelsAsciiSpaceTrimSemantics()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet[1, 1] = "";
        sheet[2, 1] = "   ";
        sheet[3, 1] = "\t";
        sheet[4, 1] = "\u00A0";
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A4")],
            Type = ConditionalFormatType.ContainsBlanks,
            Style = DifferentialStyle.Neutral,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.NotNull(evaluator.Of(At(2)).Style);
        Assert.Null(evaluator.Of(At(3)).Style);
        Assert.Null(evaluator.Of(At(4)).Style);
    }

    [Fact]
    public void DuplicatesAndUniquesAreCountedAcrossTheWholeRange()
    {
        Worksheet sheet = Column(1, 2, 2, 3);

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Duplicates("A1:A4", DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.NotNull(evaluator.Of(At(2)).Style);
        Assert.NotNull(evaluator.Of(At(3)).Style);
        Assert.Null(evaluator.Of(At(4)).Style);
    }

    [Fact]
    public void DuplicateCountingIsCaseInsensitiveAndUnderstandsDatesAndTimes()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet[1, 1] = new DateTime(2026, 1, 1);
        sheet[2, 1] = new DateTime(2026, 1, 2);
        sheet[3, 1] = new DateTime(2026, 1, 1);
        sheet[4, 1] = "Invoice";
        sheet[5, 1] = "invoice";
        sheet[6, 1] = TimeSpan.FromHours(3);
        sheet[7, 1] = TimeSpan.FromHours(3);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Duplicates("A1:A7", DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.NotNull(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
        Assert.NotNull(evaluator.Of(At(5)).Style);
        Assert.NotNull(evaluator.Of(At(6)).Style);
        Assert.NotNull(evaluator.Of(At(7)).Style);
    }

    [Theory]
    [InlineData(ConditionalFormatType.DuplicateValues)]
    [InlineData(ConditionalFormatType.UniqueValues)]
    public void DuplicateAndUniqueRulesCanonicalizeNumericText(ConditionalFormatType type)
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet[1, 1] = 1;
        sheet[2, 1] = "1.00";
        sheet[3, 1] = 2;
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A3")],
            Type = type,
            Style = DifferentialStyle.Bad,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);
        bool duplicateRule = type == ConditionalFormatType.DuplicateValues;

        Assert.Equal(duplicateRule, evaluator.Of(At(1)).Style is not null);
        Assert.Equal(duplicateRule, evaluator.Of(At(2)).Style is not null);
        Assert.Equal(!duplicateRule, evaluator.Of(At(3)).Style is not null);
    }

    [Fact]
    public void OverlappingAreasDoNotCountTheSameCellTwice()
    {
        Worksheet sheet = Column(1, 2, 3);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2"), RangeRef.Parse("A2:A3")],
            Type = ConditionalFormatType.DuplicateValues,
            Style = DifferentialStyle.Bad,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Null(evaluator.Of(At(3)).Style);
    }

    [Fact]
    public void DisjointAreasInspectOnlyTheirOccupiedCellsRatherThanTheDenseGap()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        for (int row = 1; row <= 100_000; row++)
            sheet[row, 1] = row;
        sheet[SheetLimits.MaxRow, 1] = SheetLimits.MaxRow;

        ConditionalRangeStatistics statistics = ConditionalRangeStatistics.Gather(
            sheet,
            [RangeRef.Parse("A1"), RangeRef.Parse("A1048576")],
            countOccurrences: false,
            book.DateSystem);

        Assert.Equal(2, statistics.Count);
        Assert.Equal(2, statistics.OccupiedSlotsInspected);
    }

    [Fact]
    public void TheTopValuesAreTheHighestOnesInTheRange()
    {
        Worksheet sheet = Column(10, 20, 30, 40, 50);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Top("A1:A5", 2, DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
        Assert.NotNull(evaluator.Of(At(5)).Style);
    }

    [Fact]
    public void AboveAverageComparesAgainstTheRangesMean()
    {
        Worksheet sheet = Column(1, 2, 3, 4, 5);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A5")],
            Type = ConditionalFormatType.AboveAverage,
            Above = true,
            Style = DifferentialStyle.Good,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        // The mean is 3, and "above" is strict.
        Assert.Null(evaluator.Of(At(3)).Style);
        Assert.NotNull(evaluator.Of(At(4)).Style);
    }

    [Fact]
    public void AColourScaleGradesBetweenItsStops()
    {
        Worksheet sheet = Column(0, 50, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A3", new ColorScale
        {
            Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
            Colors = [ExcelColor.FromRgb(0, 0, 0), ExcelColor.FromRgb(0xFF, 0xFF, 0xFF)],
        }));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(ExcelColor.FromRgb(0, 0, 0), evaluator.Of(At(1)).ScaleColor);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0xFF, 0xFF), evaluator.Of(At(3)).ScaleColor);

        // Halfway between black and white, give or take a rounding step.
        ExcelColor middle = evaluator.Of(At(2)).ScaleColor;
        Assert.InRange((int)(middle.Argb & 0xFF), 0x7E, 0x81);
    }

    [Fact]
    public void AColourScaleBlendsResolvedThemeColoursRatherThanSwitchingBetweenThem()
    {
        Worksheet sheet = Column(0, 50, 100);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A3", new ColorScale
        {
            Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
            Colors = [ExcelColor.FromTheme(4), ExcelColor.FromTheme(5)],
        }));

        ExcelColor middle = Evaluator(sheet).Of(At(2)).ScaleColor;

        Assert.Equal(ExcelColor.FromRgb(136, 104, 133), middle);
    }

    [Theory]
    [InlineData(ConditionalValueType.Number)]
    [InlineData(ConditionalValueType.Percent)]
    [InlineData(ConditionalValueType.Percentile)]
    [InlineData(ConditionalValueType.Formula)]
    public void AGradedStopCanTakeItsParameterFromAFormula(ConditionalValueType type)
    {
        Worksheet sheet = Column(0, 50, 100);
        ExcelColor middle = ExcelColor.FromRgb(12, 34, 56);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A3", new ColorScale
        {
            Values =
            [
                ConditionalValue.Minimum,
                new ConditionalValue(type, "HALF()"),
                ConditionalValue.Maximum,
            ],
            Colors = [ExcelColor.FromRgb(255, 0, 0), middle, ExcelColor.FromRgb(0, 255, 0)],
        }));
        var evaluatedAt = new List<CellRef>();

        ConditionalFormatEvaluator evaluator = Evaluator(sheet, (formula, cell) =>
        {
            Assert.Equal("HALF()", formula);
            evaluatedAt.Add(cell);
            return CellValue.Number(50);
        });

        _ = evaluator.Of(At(1));
        Assert.Equal(middle, evaluator.Of(At(2)).ScaleColor);
        _ = evaluator.Of(At(3));
        Assert.Equal([At(1)], evaluatedAt);
    }

    [Fact]
    public void ADataBarFillsInProportionToTheValue()
    {
        Worksheet sheet = Column(0, 25, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A3", new DataBar()));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0.1, evaluator.Of(At(1)).Bar!.Value.Fraction, 3);
        Assert.Equal(0.3, evaluator.Of(At(2)).Bar!.Value.Fraction, 3);
        Assert.Equal(0.9, evaluator.Of(At(3)).Bar!.Value.Fraction, 3);
    }

    [Fact]
    public void ADataBarHonoursItsMinimumAndMaximumLengths()
    {
        Worksheet sheet = Column(0, 100);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A2", new DataBar
        {
            MinimumLength = 20,
            MaximumLength = 60,
        }));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0.2, evaluator.Of(At(1)).Bar!.Value.Fraction, 3);
        Assert.Equal(0.6, evaluator.Of(At(2)).Bar!.Value.Fraction, 3);
    }

    [Fact]
    public void AGradedRuleWithStopIfTrueKeepsLowerRulesFromPainting()
    {
        Worksheet sheet = Column(50);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.ColorScale,
            Priority = 1,
            StopIfTrue = true,
            ColorScale = new ColorScale
            {
                Values = [ConditionalValue.Minimum, ConditionalValue.Maximum],
                Colors = [ExcelColor.FromRgb(1, 2, 3), ExcelColor.FromRgb(4, 5, 6)],
            },
        });
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["1"],
            Priority = 2,
            Style = new DifferentialStyle { Bold = true },
        });

        ConditionalAppearance appearance = Evaluator(sheet).Of(At(1));

        Assert.False(appearance.ScaleColor.IsAuto);
        Assert.Null(appearance.Style);
    }

    [Fact]
    public void AnIconSetPutsAValueInTheBandItReaches()
    {
        Worksheet sheet = Column(0, 50, 100);

        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A3", IconSet.TrafficLights));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0, evaluator.Of(At(1)).Icon!.Value.Index);
        Assert.Equal(1, evaluator.Of(At(2)).Icon!.Value.Index);
        Assert.Equal(2, evaluator.Of(At(3)).Icon!.Value.Index);
        Assert.Equal(3, evaluator.Of(At(3)).Icon!.Value.Count);
    }

    [Fact]
    public void AnIconThresholdCanExcludeTheValueExactlyOnItsBoundary()
    {
        Worksheet sheet = Column(50, 51);
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A2", new IconSet
        {
            Name = "3TrafficLights1",
            Values =
            [
                ConditionalValue.Minimum,
                new ConditionalValue(ConditionalValueType.Number, "50") { GreaterThanOrEqual = false },
            ],
        }));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Equal(0, evaluator.Of(At(1)).Icon!.Value.Index);
        Assert.Equal(1, evaluator.Of(At(2)).Icon!.Value.Index);
    }

    [Fact]
    public void AFormulaRuleWithoutAnEngineIsReportedRatherThanGuessedAt()
    {
        Worksheet sheet = Column(1);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression("A1", "$A1>0", DifferentialStyle.Good));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.Null(evaluator.Of(At(1)).Style);
        Assert.Contains("conditional-format-formula", evaluator.Unsupported);
    }

    [Fact]
    public void AFormulaRuleWithAnEnginePaints()
    {
        Worksheet sheet = Column(1, 0);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.Expression("A1:A2", "$A1>0", DifferentialStyle.Good));

        // Stands in for the formula package: the rule is true where the cell holds something.
        ConditionalFormatEvaluator evaluator = Evaluator(
            sheet,
            (formula, cell) => CellValue.Boolean(sheet.Values.Get(cell.Row, cell.Column).ToSerialNumber() > 0));

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Empty(evaluator.Unsupported);
    }

    [Fact]
    public void AFormulaIsTranslatedFromTheFirstRangesTopLeftCell()
    {
        Worksheet sheet = Column(1);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("B2:B3"), RangeRef.Parse("D5")],
            Type = ConditionalFormatType.Expression,
            Formulas = ["A2+$A2+A$2+$A$2>0"],
            Style = DifferentialStyle.Good,
        });
        var formulas = new List<string>();
        ConditionalFormatEvaluator evaluator = Evaluator(sheet, (formula, _) =>
        {
            formulas.Add(formula);
            return CellValue.Boolean(true);
        });

        evaluator.Of(CellRef.Parse("B2"));
        evaluator.Of(CellRef.Parse("D5"));

        Assert.Equal(["A2+$A2+A$2+$A$2>0", "C5+$A5+C$2+$A$2>0"], formulas);
    }

    [Fact]
    public void ACopyAwareEvaluatorReceivesStoredTextAnchorAndTargetSeparately()
    {
        Worksheet sheet = Column(1);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("B2:B3")],
            Type = ConditionalFormatType.Expression,
            Formulas = ["$A2>0"],
            Style = DifferentialStyle.Good,
        });
        var calls = new List<(string Formula, CellRef Anchor, CellRef Target)>();
        var evaluator = new ConditionalFormatEvaluator(
            sheet,
            evaluateRelative: (formula, anchor, target) =>
            {
                calls.Add((formula, anchor, target));
                return CellValue.Boolean(true);
            });

        evaluator.Of(CellRef.Parse("B2"));
        evaluator.Of(CellRef.Parse("B3"));

        Assert.Equal(
            [
                ("$A2>0", CellRef.Parse("B2"), CellRef.Parse("B2")),
                ("$A2>0", CellRef.Parse("B2"), CellRef.Parse("B3")),
            ],
            calls);
    }

    [Fact]
    public void AComparisonFormulaIsTranslatedForEachCell()
    {
        Worksheet sheet = Column(10, 10);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Formulas = ["$B1"],
            Style = DifferentialStyle.Good,
        });
        var formulas = new List<string>();
        ConditionalFormatEvaluator evaluator = Evaluator(sheet, (formula, _) =>
        {
            formulas.Add(formula);
            return CellValue.Number(formula == "$B1" ? 5 : 15);
        });

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
        Assert.Equal(["$B1", "$B2"], formulas);
    }

    [Theory]
    [InlineData(ConditionalFormatType.ColorScale)]
    [InlineData(ConditionalFormatType.DataBar)]
    [InlineData(ConditionalFormatType.IconSet)]
    public void AGradedRuleAppliesItsActivityFormulaFirst(ConditionalFormatType type)
    {
        Worksheet sheet = Column(0, 100);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2")],
            Type = type,
            Formulas = ["$B1"],
            ColorScale = type == ConditionalFormatType.ColorScale ? ColorScale.RedYellowGreen : null,
            DataBar = type == ConditionalFormatType.DataBar ? new DataBar() : null,
            IconSet = type == ConditionalFormatType.IconSet ? IconSet.TrafficLights : null,
        });
        var formulas = new List<string>();
        ConditionalFormatEvaluator evaluator = Evaluator(sheet, (formula, _) =>
        {
            formulas.Add(formula);
            return CellValue.Boolean(formula == "$B2");
        });

        ConditionalAppearance inactive = evaluator.Of(At(1));
        ConditionalAppearance active = evaluator.Of(At(2));

        Assert.True(inactive.IsEmpty);
        Assert.False(active.IsEmpty);
        Assert.Equal(["$B1", "$B2"], formulas);
    }

    [Fact]
    public void AnInactiveGradedRuleDoesNotStopLowerRules()
    {
        Worksheet sheet = Column(100);
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.ColorScale,
            Priority = 1,
            StopIfTrue = true,
            Formulas = ["FALSE"],
            ColorScale = ColorScale.RedYellowGreen,
        });
        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.GreaterThan,
            Priority = 2,
            Formulas = ["0"],
            Style = DifferentialStyle.Good,
        });

        ConditionalAppearance appearance = Evaluator(
            sheet,
            (formula, _) => CellValue.Boolean(!formula.Equals("FALSE", StringComparison.OrdinalIgnoreCase)))
            .Of(At(1));

        Assert.Equal(DifferentialStyle.Good, appearance.Style);
        Assert.True(appearance.ScaleColor.IsAuto);
    }

    [Fact]
    public void ATimePeriodRuleComparesAgainstTheGivenToday()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = Today.ToDateTime(TimeOnly.MinValue);
        sheet["A2"] = Today.AddDays(-1).ToDateTime(TimeOnly.MinValue);

        sheet.ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A2")],
            Type = ConditionalFormatType.TimePeriod,
            TimePeriod = "today",
            Style = DifferentialStyle.Neutral,
        });

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.Null(evaluator.Of(At(2)).Style);
    }

    [Fact]
    public void ACellOutsideEveryRangeIsLeftAlone()
    {
        Worksheet sheet = Column(500, 500);
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1", 1, DifferentialStyle.Bad));

        ConditionalFormatEvaluator evaluator = Evaluator(sheet);

        Assert.NotNull(evaluator.Of(At(1)).Style);
        Assert.True(evaluator.Of(At(2)).IsEmpty);
    }

    [Fact]
    public void ASheetWithoutRulesHasNothingToApply()
    {
        Assert.False(Evaluator(Column(1)).HasRules);
    }
}
