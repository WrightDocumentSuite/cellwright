using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Cellwright.Streaming;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Conditional formatting from both ends: rules built in the model have to come back after a
/// round trip and be understood by another library, and rules read from a file have to survive
/// even when the model has no way to describe them.
/// </summary>
public sealed class ConditionalFormatTests
{
    [Fact]
    public void ConditionalValueKeepsItsTwoPartPositionalApi()
    {
        var stop = new ConditionalValue(ConditionalValueType.Percent, "25")
        {
            GreaterThanOrEqual = false,
        };

        (ConditionalValueType type, string? value) = stop;

        Assert.Equal(ConditionalValueType.Percent, type);
        Assert.Equal("25", value);
        Assert.False(stop.GreaterThanOrEqual);
        Assert.True(default(ConditionalValue).GreaterThanOrEqual);
        Assert.Equal(default, ConditionalValue.Minimum);
    }

    [Fact]
    public void AddingADuplicatePriorityAssignsTheNextFreePriority()
    {
        Workbook book = Sample();
        ConditionalFormatRule first = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 1, DifferentialStyle.Good);
        first.Priority = 7;
        ConditionalFormatRule second = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 2, DifferentialStyle.Bad);
        second.Priority = 7;

        book.Sheets[0].ConditionalFormats.Add(first);
        book.Sheets[0].ConditionalFormats.Add(second);

        Assert.Equal(7, first.Priority);
        Assert.Equal(8, second.Priority);
    }

    [Fact]
    public void AddingAfterTheMaximumPriorityUsesTheFirstPositiveGap()
    {
        Workbook book = Sample();
        ConditionalFormatRule first = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 1, DifferentialStyle.Good);
        first.Priority = int.MaxValue;

        book.Sheets[0].ConditionalFormats.Add(first);
        ConditionalFormatRule second = book.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 2, DifferentialStyle.Bad));

        Assert.Equal(int.MaxValue, first.Priority);
        Assert.Equal(1, second.Priority);
    }

    [Fact]
    public async Task AnOwnedRuleCannotBeReusedAndItsOriginalSheetKeepsTrackingIt()
    {
        Workbook source = Sample();
        source.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Good));
        Workbook loaded = await SaveAndReloadAsync(source);
        ConditionalFormatRule rule = Assert.Single(loaded.Sheets[0].ConditionalFormats);
        Worksheet other = loaded.Sheets.Add("Other");
        other.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 100, DifferentialStyle.Bad));

        ArgumentException error = Assert.Throws<ArgumentException>(
            () => other.ConditionalFormats.Add(rule));

        Assert.Equal("rule", error.ParamName);
        Assert.Equal(1, rule.Priority);
        Assert.Single(loaded.Sheets[0].ConditionalFormats);
        Assert.Single(other.ConditionalFormats);

        rule.Priority = 17;
        Workbook regenerated = await SaveAndReloadAsync(loaded);

        Assert.Equal(17, Assert.Single(regenerated.Sheets[0].ConditionalFormats).Priority);
    }

    [Fact]
    public void RemovingARuleReleasesItForAnotherCollection()
    {
        Workbook book = Sample();
        Worksheet other = book.Sheets.Add("Other");
        ConditionalFormatRule rule = book.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Good));

        Assert.True(book.Sheets[0].ConditionalFormats.Remove(rule));
        Assert.Same(rule, other.ConditionalFormats.Add(rule));
        Assert.Empty(book.Sheets[0].ConditionalFormats);
        Assert.Same(rule, Assert.Single(other.ConditionalFormats));
    }

    [Fact]
    public void ClearingRulesReleasesEveryRuleForAnotherCollection()
    {
        Workbook book = Sample();
        Worksheet source = book.Sheets[0];
        Worksheet other = book.Sheets.Add("Other");
        ConditionalFormatRule first = source.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Good));
        ConditionalFormatRule second = source.ConditionalFormats.Add(
            ConditionalFormatRule.CellLessThan("A1:A10", 200, DifferentialStyle.Bad));

        source.ConditionalFormats.Clear();
        other.ConditionalFormats.Add(first);
        other.ConditionalFormats.Add(second);

        Assert.Empty(source.ConditionalFormats);
        Assert.Equal(new[] { first, second }, other.ConditionalFormats.ToArray());
    }

    [Fact]
    public void FileLoadingCannotAttachTheSameRuleToTwoCollections()
    {
        Workbook book = Sample();
        Worksheet other = book.Sheets.Add("Other");
        ConditionalFormatRule rule = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 500, DifferentialStyle.Good);
        book.Sheets[0].ConditionalFormats.AddFromFile(rule);

        ArgumentException error = Assert.Throws<ArgumentException>(
            () => other.ConditionalFormats.AddFromFile(rule));

        Assert.Equal("rule", error.ParamName);
        Assert.Single(book.Sheets[0].ConditionalFormats);
        Assert.Empty(other.ConditionalFormats);
    }

    [Fact]
    public async Task RegeneratedMarkupPreservesUniquePrioritiesAndRepairsOnlyDuplicates()
    {
        Workbook book = Sample();
        ConditionalFormatRule first = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 1, DifferentialStyle.Good);
        first.Priority = 5;
        ConditionalFormatRule second = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 2, DifferentialStyle.Bad);
        second.Priority = 6;
        ConditionalFormatRule third = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 3, DifferentialStyle.Neutral);
        third.Priority = 10;

        book.Sheets[0].ConditionalFormats.Add(first);
        book.Sheets[0].ConditionalFormats.Add(second);
        book.Sheets[0].ConditionalFormats.Add(third);
        second.Priority = 5;

        Workbook reloaded = await SaveAndReloadAsync(book);

        Assert.Equal(
            [5, 6, 10],
            reloaded.Sheets[0].ConditionalFormats.Select(static rule => rule.Priority));
    }

    [Fact]
    public async Task ExtensionOnlyRulesReserveTheirWorksheetWidePriorities()
    {
        using MemoryStream ordinary = await SaveAsync(Sample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, X14ConditionalFormats);
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        Worksheet sheet = loaded.Sheets[0];

        Assert.Empty(sheet.ConditionalFormats);
        Assert.Contains(
            "conditional-format-opaque",
            new ConditionalFormatEvaluator(sheet).Unsupported);

        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("C1:C10", 500, DifferentialStyle.Good));

        using MemoryStream saved = await SaveAsync(loaded);
        (int[] basePriorities, int[] extensionPriorities) = ReadPriorities(saved);

        Assert.Equal([3], basePriorities);
        Assert.Equal([1, 2], extensionPriorities);
        Assert.Equal(
            3,
            basePriorities.Concat(extensionPriorities).Distinct().Count());
    }

    [Fact]
    public async Task SupportedExtensionIconSetsAreProjectedWithoutBecomingEditableRules()
    {
        using MemoryStream ordinary = await SaveAsync(IconSample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, SupportedX14IconSets);
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        Worksheet sheet = loaded.Sheets[0];
        var evaluator = new ConditionalFormatEvaluator(sheet);

        Assert.Empty(sheet.ConditionalFormats);
        Assert.True(evaluator.HasRules);
        Assert.Empty(evaluator.Unsupported);

        ConditionalIcon triangle = Assert.IsType<ConditionalIcon>(
            evaluator.Of(CellRef.Parse("A1")).Icon);
        Assert.Equal("3Triangles", triangle.SetName);
        Assert.Equal(0, triangle.Index);

        // The exact 50% boundary is exclusive. It therefore remains in the lowest band, which
        // reverse maps to the last icon; the rule also hides the cell value.
        ConditionalIcon star = Assert.IsType<ConditionalIcon>(
            evaluator.Of(CellRef.Parse("B2")).Icon);
        Assert.Equal("3Stars", star.SetName);
        Assert.Equal(2, star.Index);
        Assert.False(star.ShowValue);

        ConditionalIcon boxes = Assert.IsType<ConditionalIcon>(
            evaluator.Of(CellRef.Parse("C2")).Icon);
        Assert.Equal("5Boxes", boxes.SetName);
        Assert.Equal(2, boxes.Index);
        Assert.Equal(5, boxes.Count);
    }

    [Fact]
    public async Task ExtensionAndBaseIconSetsShareOnePriorityOrder()
    {
        Workbook source = IconSample();
        ConditionalFormatRule ordinaryRule = ConditionalFormatRule.Icons(
            "A1:A3", IconSet.TrafficLights);
        ordinaryRule.Priority = 2;
        source.Sheets[0].ConditionalFormats.Add(ordinaryRule);

        using MemoryStream ordinary = await SaveAsync(source);
        using MemoryStream package = AppendWorksheetExtension(ordinary, PriorityOneX14Stars);
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        var evaluator = new ConditionalFormatEvaluator(loaded.Sheets[0]);
        ConditionalIcon icon = Assert.IsType<ConditionalIcon>(
            evaluator.Of(CellRef.Parse("A1")).Icon);

        Assert.Equal("3Stars", icon.SetName);
        Assert.Empty(evaluator.Unsupported);
    }

    [Fact]
    public async Task CustomExtensionIconSetStaysOpaqueRatherThanBeingApproximated()
    {
        using MemoryStream ordinary = await SaveAsync(IconSample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, CustomX14IconSet);
        string before = ExtensionOf(ReadEntry(package, "xl/worksheets/sheet1.xml"));
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        var evaluator = new ConditionalFormatEvaluator(loaded.Sheets[0]);

        Assert.Empty(loaded.Sheets[0].ConditionalFormats);
        Assert.False(evaluator.HasRules);
        Assert.Contains("conditional-format-opaque", evaluator.Unsupported);

        using MemoryStream saved = await SaveAsync(loaded);
        string after = ExtensionOf(ReadEntry(saved, "xl/worksheets/sheet1.xml"));
        Assert.True(XNode.DeepEquals(XElement.Parse(before), XElement.Parse(after)));
    }

    [Fact]
    public async Task ProjectingAnExtensionDoesNotChangeItsSavedXml()
    {
        using MemoryStream ordinary = await SaveAsync(IconSample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, SupportedX14IconSets);
        string before = ExtensionOf(ReadEntry(package, "xl/worksheets/sheet1.xml"));
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        loaded.Sheets[0][1, 1] = 1234;
        using MemoryStream saved = await SaveAsync(loaded);
        string after = ExtensionOf(ReadEntry(saved, "xl/worksheets/sheet1.xml"));

        Assert.True(
            XNode.DeepEquals(XElement.Parse(before), XElement.Parse(after)),
            $"The preserved x14 extension changed.\nBefore: {before}\nAfter: {after}");
    }

    [Fact]
    public async Task StructuralEditDropsOnlyX14ConditionalFormattingAndWarns()
    {
        string extension = SupportedX14IconSets.Replace(
            "</extLst>",
            """
              <ext uri="{00000000-0000-0000-0000-000000000099}">
                <probe xmlns="urn:cellwright:test">kept</probe>
              </ext>
            </extLst>
            """,
            StringComparison.Ordinal);
        using MemoryStream ordinary = await SaveAsync(IconSample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, extension);
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        var warnings = new List<WorkbookWarning>();
        loaded.OnWarning = warnings.Add;

        loaded.Sheets[0].InsertRows(1);

        Assert.Contains(
            warnings,
            static warning => warning.Code == "StructuralEditDroppedExtensionConditionalFormats");
        var evaluator = new ConditionalFormatEvaluator(loaded.Sheets[0]);
        Assert.False(evaluator.HasRules);
        Assert.Empty(evaluator.Unsupported);
        Assert.Null(loaded.Sheets[0].Markup?.ExtensionConditionalFormatPriorities);

        using MemoryStream saved = await SaveAsync(loaded);
        string worksheet = ReadEntry(saved, "xl/worksheets/sheet1.xml");
        XDocument xml = XDocument.Parse(worksheet);
        XNamespace x14 = X14Namespace;
        XNamespace probe = "urn:cellwright:test";

        Assert.Empty(xml.Descendants(x14 + "conditionalFormattings"));
        Assert.Equal("kept", Assert.Single(xml.Descendants(probe + "probe")).Value);
        Assert.DoesNotContain("3Stars", worksheet, StringComparison.Ordinal);
        Assert.DoesNotContain("3Triangles", worksheet, StringComparison.Ordinal);
        Assert.DoesNotContain("5Boxes", worksheet, StringComparison.Ordinal);

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        Assert.False(new ConditionalFormatEvaluator(reloaded.Sheets[0]).HasRules);
    }

    [Fact]
    public async Task StructuralEditKeepsX14ConditionalFormattingOnOtherSheets()
    {
        Workbook source = IconSample();
        source.Sheets.Add("Other");
        using MemoryStream ordinary = await SaveAsync(source);
        using MemoryStream package = AppendWorksheetExtension(ordinary, PriorityOneX14Stars);
        string before = ExtensionOf(ReadEntry(package, "xl/worksheets/sheet1.xml"));
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        var warnings = new List<WorkbookWarning>();
        loaded.OnWarning = warnings.Add;

        loaded.Sheets["Other"].InsertRows(1);

        Assert.DoesNotContain(
            warnings,
            static warning => warning.Code == "StructuralEditDroppedExtensionConditionalFormats");
        Assert.True(new ConditionalFormatEvaluator(loaded.Sheets[0]).HasRules);
        using MemoryStream saved = await SaveAsync(loaded);
        string after = ExtensionOf(ReadEntry(saved, "xl/worksheets/sheet1.xml"));
        Assert.True(XNode.DeepEquals(XElement.Parse(before), XElement.Parse(after)));
    }

    [Fact]
    public async Task SheetRenameKeepsX14ConditionalFormatting()
    {
        using MemoryStream ordinary = await SaveAsync(IconSample());
        using MemoryStream package = AppendWorksheetExtension(ordinary, PriorityOneX14Stars);
        string before = ExtensionOf(ReadEntry(package, "xl/worksheets/sheet1.xml"));
        Workbook loaded = await Workbook.LoadAsync(
            package,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        var warnings = new List<WorkbookWarning>();
        loaded.OnWarning = warnings.Add;

        loaded.Sheets[0].Name = "Renamed";

        Assert.DoesNotContain(
            warnings,
            static warning => warning.Code.Contains(
                "DroppedExtensionConditionalFormats",
                StringComparison.Ordinal));
        Assert.True(new ConditionalFormatEvaluator(loaded.Sheets[0]).HasRules);
        using MemoryStream saved = await SaveAsync(loaded);
        string after = ExtensionOf(ReadEntry(saved, "xl/worksheets/sheet1.xml"));
        Assert.True(XNode.DeepEquals(XElement.Parse(before), XElement.Parse(after)));
    }

    [Theory]
    [InlineData("0", false)]
    [InlineData("1", true)]
    [InlineData("false", false)]
    [InlineData("true", true)]
    public void ExtensionProjectionAcceptsXsdBooleanLexicalForms(string text, bool expected)
    {
        string extension = PriorityOneX14Stars.Replace(
            "iconSet=\"3Stars\"",
            $"iconSet=\"3Stars\" reverse=\"{text}\"",
            StringComparison.Ordinal);
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            ConditionalFormatPart.ReadExtensionConditionalFormats(
                extension,
                out _,
                out _));

        Assert.False(projection.HasUnsupportedRules);
        Assert.Equal(expected, Assert.Single(projection.Rules).IconSet!.Reverse);
    }

    [Theory]
    [InlineData("TRUE")]
    [InlineData("FALSE")]
    public void ExtensionProjectionRejectsNonXsdBooleanLexicalForms(string text)
    {
        string extension = PriorityOneX14Stars.Replace(
            "iconSet=\"3Stars\"",
            $"iconSet=\"3Stars\" reverse=\"{text}\"",
            StringComparison.Ordinal);
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            ConditionalFormatPart.ReadExtensionConditionalFormats(
                extension,
                out _,
                out _));

        Assert.Empty(projection.Rules);
        Assert.True(projection.HasUnsupportedRules);
    }

    [Theory]
    [InlineData("min")]
    [InlineData("max")]
    public void ExtensionProjectionRejectsMinAndMaxIconThresholds(string type)
    {
        string extension = PriorityOneX14Stars.Replace(
            "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo>",
            $"<x14:cfvo type=\"{type}\"/>",
            StringComparison.Ordinal);
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            ConditionalFormatPart.ReadExtensionConditionalFormats(
                extension,
                out _,
                out _));

        Assert.Empty(projection.Rules);
        Assert.True(projection.HasUnsupportedRules);
    }

    [Theory]
    [InlineData("3Arrows", 3)]
    [InlineData("NoIcons", 3)]
    [InlineData("3Stars", 2)]
    [InlineData("5Boxes", 4)]
    public void ExtensionProjectionRejectsRulesOutsideTheExactSupportedSubset(
        string iconSet,
        int thresholdCount)
    {
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            ConditionalFormatPart.ReadExtensionConditionalFormats(
                X14IconSetExtension(iconSet, thresholdCount),
                out IReadOnlySet<int>? priorities,
                out bool hasConditionalFormats));

        Assert.True(hasConditionalFormats);
        Assert.Equal([7], priorities!.Order());
        Assert.Empty(projection.Rules);
        Assert.True(projection.HasUnsupportedRules);
    }

    [Fact]
    public void ExtensionCompanionWithoutPriorityDoesNotReserveANumber()
    {
        IReadOnlySet<int>? priorities = ConditionalFormatPart.ReadExtensionPriorities(
            $"""
             <extLst xmlns:x14="{X14Namespace}">
               <x14:conditionalFormatting>
                 <x14:cfRule type="dataBar"/>
               </x14:conditionalFormatting>
             </extLst>
             """,
            out bool hasConditionalFormats);

        Assert.True(hasConditionalFormats);
        Assert.Null(priorities);
    }

    [Fact]
    public void ExtensionPriorityAcceptsTheXsdIntLexicalForm()
    {
        IReadOnlySet<int>? priorities = ConditionalFormatPart.ReadExtensionPriorities(
            $"<extLst xmlns:x14=\"{X14Namespace}\"><x14:cfRule priority=\" +7 \"/></extLst>",
            out bool hasConditionalFormats);

        Assert.True(hasConditionalFormats);
        Assert.NotNull(priorities);
        Assert.Equal([7], priorities.Order());
    }

    [Fact]
    public async Task BaseRuleIntegersAcceptTheXsdWhitespaceAndLeadingSign()
    {
        Workbook source = Sample();
        Worksheet sheet = source.Sheets[0];

        ConditionalFormatRule lower = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 0, new DifferentialStyle { FillColor = ExcelColor.FromRgb(0xFF, 0x00, 0x00) });
        lower.Priority = 7;
        ConditionalFormatRule higher = ConditionalFormatRule.CellGreaterThan(
            "A1:A10", 0, new DifferentialStyle { FillColor = ExcelColor.FromRgb(0x00, 0xFF, 0x00) });
        higher.Priority = 2;
        ConditionalFormatRule bar = ConditionalFormatRule.Bars("A1:A10", new DataBar
        {
            MinimumLength = 15,
            MaximumLength = 85,
        });
        bar.Priority = 9;
        ConditionalFormatRule top = ConditionalFormatRule.Top("A1:A10", 10, DifferentialStyle.Neutral);
        top.Priority = 10;

        sheet.ConditionalFormats.Add(lower);
        sheet.ConditionalFormats.Add(higher);
        sheet.ConditionalFormats.Add(bar);
        sheet.ConditionalFormats.Add(top);

        using MemoryStream ordinary = await SaveAsync(source);
        using MemoryStream lexical = RewriteWorksheet(ordinary, static xml => xml
            .Replace("priority=\"7\"", "priority=\" +7 \"", StringComparison.Ordinal)
            .Replace("minLength=\"15\"", "minLength=\" 15 \"", StringComparison.Ordinal)
            .Replace("rank=\"10\"", "rank=\" 10 \"", StringComparison.Ordinal));
        Workbook loaded = await Workbook.LoadAsync(
            lexical,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        Assert.Equal([7, 2, 9, 10], loaded.Sheets[0].ConditionalFormats.Select(static rule => rule.Priority));
        Assert.Equal(
            ExcelColor.FromRgb(0x00, 0xFF, 0x00),
            new ConditionalFormatEvaluator(loaded.Sheets[0]).Of(CellRef.Parse("A1")).Style!.FillColor);
        DataBar loadedBar = Assert.IsType<DataBar>(loaded.Sheets[0].ConditionalFormats[2].DataBar);
        Assert.Equal(15, loadedBar.MinimumLength);
        Assert.Equal(85, loadedBar.MaximumLength);
        Assert.Equal(10, loaded.Sheets[0].ConditionalFormats[3].Rank);

        // Any tracked edit regenerates all base rules, so the parsed values must survive that path.
        loaded.Sheets[0].ConditionalFormats[0].Ranges = [.. loaded.Sheets[0].ConditionalFormats[0].Ranges];
        Workbook regenerated = await SaveAndReloadAsync(loaded);

        Assert.Equal(
            [7, 2, 9, 10],
            regenerated.Sheets[0].ConditionalFormats.Select(static rule => rule.Priority));
        DataBar regeneratedBar = Assert.IsType<DataBar>(regenerated.Sheets[0].ConditionalFormats[2].DataBar);
        Assert.Equal(15, regeneratedBar.MinimumLength);
        Assert.Equal(85, regeneratedBar.MaximumLength);
        Assert.Equal(10, regenerated.Sheets[0].ConditionalFormats[3].Rank);
    }

    [Fact]
    public void MalformedExtensionMarkupDoesNotAbortPriorityDiscovery()
    {
        IReadOnlySet<int>? priorities = ConditionalFormatPart.ReadExtensionPriorities(
            $"<extLst xmlns:x14=\"{X14Namespace}\"><x14:conditionalFormatting>" +
            "<x14:cfRule priority=\"7\"/><",
            out bool hasConditionalFormats);

        Assert.True(hasConditionalFormats);
        Assert.NotNull(priorities);
        Assert.Equal([7], priorities.Order());
    }

    [Theory]
    [InlineData("priority")]
    [InlineData("ranges")]
    [InlineData("formulas")]
    public async Task EditingALoadedRuleRegeneratesItsMarkup(string property)
    {
        Workbook source = Sample();
        source.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Good));
        Workbook loaded = await SaveAndReloadAsync(source);
        ConditionalFormatRule rule = Assert.Single(loaded.Sheets[0].ConditionalFormats);

        switch (property)
        {
            case "priority":
                rule.Priority = 17;
                break;
            case "ranges":
                rule.Ranges = [RangeRef.Parse("B2:B8")];
                break;
            default:
                rule.Formulas = ["750"];
                break;
        }

        ConditionalFormatRule saved = Assert.Single(
            (await SaveAndReloadAsync(loaded)).Sheets[0].ConditionalFormats);

        Assert.Equal(property == "priority" ? 17 : 1, saved.Priority);
        Assert.Equal(
            RangeRef.Parse(property == "ranges" ? "B2:B8" : "A1:A10"),
            Assert.Single(saved.Ranges));
        Assert.Equal([property == "formulas" ? "750" : "500"], saved.Formulas);
    }

    [Fact]
    public async Task OpaquePriorityRewriteAcceptsXmlWhitespaceAndAddsAMissingAttribute()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        sheet["A1"] = 1;
        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A10")],
            Priority = 11,
            RawXml = "<cfRule\n type='uniqueValues'\n priority = '7'/>",
        });
        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("B1:B10")],
            Priority = 12,
            RawXml = "<cfRule type='duplicateValues'/>",
        });
        sheet.ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("C1:C10", 1, DifferentialStyle.Good));

        using MemoryStream stream = await SaveAsync(book);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        WorksheetPart part = Assert.Single(document.WorkbookPart!.WorksheetParts);
        using var reader = new StreamReader(part.GetStream());
        string xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);

        Assert.Contains("priority = '11'", xml, StringComparison.Ordinal);
        Assert.Contains("<cfRule type='duplicateValues' priority=\"12\"/>", xml, StringComparison.Ordinal);
    }

    private static async Task<MemoryStream> SaveAsync(Workbook workbook)
    {
        var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        return stream;
    }

    private static (int[] Base, int[] Extension) ReadPriorities(MemoryStream package)
    {
        string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
        var basePriorities = new List<int>();
        var extensionPriorities = new List<int>();
        using var xml = XmlReader.Create(new StringReader(sheet), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        });
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "cfRule" ||
                !int.TryParse(xml.GetAttribute("priority"), out int priority))
            {
                continue;
            }

            if (xml.NamespaceURI == X14Namespace)
                extensionPriorities.Add(priority);
            else if (xml.NamespaceURI == SpreadsheetNamespace)
                basePriorities.Add(priority);
        }

        return ([.. basePriorities], [.. extensionPriorities]);
    }

    private static MemoryStream AppendWorksheetExtension(MemoryStream package, string extension)
    {
        return RewriteWorksheet(package, sheet => sheet.Replace(
            "</worksheet>", extension + "</worksheet>", StringComparison.Ordinal));
    }

    private static MemoryStream RewriteWorksheet(MemoryStream package, Func<string, string> rewrite)
    {
        string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
        string rewrittenSheet = rewrite(sheet);
        Assert.NotEqual(sheet, rewrittenSheet);

        var rewritten = new MemoryStream();
        package.Position = 0;
        using (var source = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                using Stream output = target.CreateEntry(entry.FullName).Open();
                if (entry.FullName == "xl/worksheets/sheet1.xml")
                {
                    output.Write(Encoding.UTF8.GetBytes(rewrittenSheet));
                }
                else
                {
                    using Stream input = entry.Open();
                    input.CopyTo(output);
                }
            }
        }

        rewritten.Position = 0;
        return rewritten;
    }

    private static string ReadEntry(MemoryStream package, string name)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = Assert.IsType<ZipArchiveEntry>(archive.GetEntry(name));
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static string ExtensionOf(string worksheet)
    {
        using var xml = XmlReader.Create(new StringReader(worksheet), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        });
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "extLst")
                return xml.ReadOuterXml();
        }

        return string.Empty;
    }

    private static string X14IconSetExtension(string iconSet, int thresholdCount)
    {
        string thresholds = string.Join(
            Environment.NewLine,
            Enumerable.Range(0, thresholdCount).Select(index =>
                $"<x14:cfvo type=\"percent\"><xm:f>{index}</xm:f></x14:cfvo>"));
        return $$"""
            <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                   xmlns:x14="{{X14Namespace}}"
                   xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                <x14:conditionalFormattings>
                  <x14:conditionalFormatting>
                    <x14:cfRule type="iconSet" priority="7">
                      <x14:iconSet iconSet="{{iconSet}}">
                        {{thresholds}}
                      </x14:iconSet>
                    </x14:cfRule>
                    <xm:sqref>A1:A3</xm:sqref>
                  </x14:conditionalFormatting>
                </x14:conditionalFormattings>
              </ext>
            </extLst>
            """;
    }

    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string X14Namespace =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string X14ConditionalFormats = """
        <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
               xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
               xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:conditionalFormattings>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="1" id="{00000000-0000-0000-0000-000000000001}">
                  <x14:iconSet iconSet="3Stars">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>A1:A10</xm:sqref>
              </x14:conditionalFormatting>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="2" id="{00000000-0000-0000-0000-000000000002}">
                  <x14:iconSet iconSet="3Triangles">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>B1:B10</xm:sqref>
              </x14:conditionalFormatting>
              <x14:conditionalFormatting>
                <x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000003}"/>
                <xm:sqref>D1:D10</xm:sqref>
              </x14:conditionalFormatting>
            </x14:conditionalFormattings>
          </ext>
        </extLst>
        """;

    private const string SupportedX14IconSets = """
        <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
               xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
               xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:conditionalFormattings>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="3">
                  <x14:iconSet iconSet="3Triangles">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>50</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>75</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>A1:A3</xm:sqref>
              </x14:conditionalFormatting>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="1">
                  <x14:iconSet iconSet="3Stars" reverse="1" showValue="0">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent" gte="0"><xm:f>50</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>75</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>B1:B3</xm:sqref>
              </x14:conditionalFormatting>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="2">
                  <x14:iconSet iconSet="5Boxes">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>20</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>40</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>60</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>80</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>C1:C5</xm:sqref>
              </x14:conditionalFormatting>
            </x14:conditionalFormattings>
          </ext>
        </extLst>
        """;

    private const string PriorityOneX14Stars = """
        <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
               xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
               xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:conditionalFormattings>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="1">
                  <x14:iconSet iconSet="3Stars">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>A1:A3</xm:sqref>
              </x14:conditionalFormatting>
            </x14:conditionalFormattings>
          </ext>
        </extLst>
        """;

    private const string CustomX14IconSet = """
        <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
               xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
               xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:conditionalFormattings>
              <x14:conditionalFormatting>
                <x14:cfRule type="iconSet" priority="1">
                  <x14:iconSet iconSet="3Arrows" custom="1">
                    <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
                    <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
                    <x14:cfIcon iconSet="3Arrows" iconId="0"/>
                    <x14:cfIcon iconSet="NoIcons" iconId="0"/>
                    <x14:cfIcon iconSet="3Arrows" iconId="2"/>
                  </x14:iconSet>
                </x14:cfRule>
                <xm:sqref>A1:A3</xm:sqref>
              </x14:conditionalFormatting>
            </x14:conditionalFormattings>
          </ext>
        </extLst>
        """;

    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using MemoryStream stream = await SaveAsync(workbook);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        for (int row = 1; row <= 10; row++)
        {
            sheet[row, 1] = row * 100;
            sheet[row, 2] = $"item-{row}";
        }

        return book;
    }

    private static Workbook IconSample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        int[] values = [0, 50, 100, 75, 25];
        for (int row = 1; row <= values.Length; row++)
        {
            for (int column = 1; column <= 3; column++)
                sheet[row, column] = values[row - 1];
        }

        return book;
    }

    [Fact]
    public async Task EveryRuleKindRoundTrips()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellBetween("A1:A10", 200, 400, DifferentialStyle.Neutral));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Expression("B1:B10", "$A1>800", DifferentialStyle.Good));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Containing("B1:B10", "item-3", DifferentialStyle.Bad));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Top("A1:A10", 3, DifferentialStyle.Good));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Duplicates("B1:B10", DifferentialStyle.Neutral));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A10", ColorScale.RedYellowGreen));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A10", new DataBar()));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A10", IconSet.TrafficLights));

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(9, loaded.ConditionalFormats.Count);
        Assert.Equal(
            [
                ConditionalFormatType.CellIs, ConditionalFormatType.CellIs, ConditionalFormatType.Expression,
                ConditionalFormatType.ContainsText, ConditionalFormatType.Top10, ConditionalFormatType.DuplicateValues,
                ConditionalFormatType.ColorScale, ConditionalFormatType.DataBar, ConditionalFormatType.IconSet,
            ],
            loaded.ConditionalFormats.Select(static rule => rule.Type));
        Assert.Equal(Enumerable.Range(1, 9), loaded.ConditionalFormats.Select(static rule => rule.Priority));
    }

    [Fact]
    public async Task ComparisonRuleKeepsItsOperatorBoundsAndRanges()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.CellBetween("A1:A10", 200, 400, DifferentialStyle.Bad));

        ConditionalFormatRule rule = Assert.Single((await SaveAndReloadAsync(book)).Sheets[0].ConditionalFormats);

        Assert.Equal(ConditionalOperator.Between, rule.Operator);
        Assert.Equal(["200", "400"], rule.Formulas);
        Assert.Equal(RangeRef.Parse("A1:A10"), Assert.Single(rule.Ranges));
    }

    [Fact]
    public async Task GradedRulesKeepTheirStopsAndColors()
    {
        Workbook book = Sample();
        Worksheet sheet = book.Sheets[0];
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A10", ColorScale.RedYellowGreen));
        sheet.ConditionalFormats.Add(ConditionalFormatRule.Bars("B1:B10", new DataBar
        {
            Minimum = ConditionalValue.Number(0),
            Maximum = ConditionalValue.Number(1000),
            Color = ExcelColor.FromRgb(0x63, 0x8E, 0xC6),
            ShowValue = false,
        }));

        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets[0];

        ColorScale scale = Assert.IsType<ColorScale>(loaded.ConditionalFormats[0].ColorScale);
        Assert.Equal(3, scale.Values.Count);
        Assert.Equal(ConditionalValueType.Minimum, scale.Values[0].Type);
        Assert.Equal(ConditionalValueType.Percentile, scale.Values[1].Type);
        Assert.Equal("50", scale.Values[1].Value);
        Assert.Equal(ExcelColor.FromRgb(0x63, 0xBE, 0x7B), scale.Colors[2]);

        DataBar bar = Assert.IsType<DataBar>(loaded.ConditionalFormats[1].DataBar);
        Assert.False(bar.ShowValue);
        Assert.Equal("1000", bar.Maximum.Value);
    }

    [Fact]
    public async Task FormulaValuedPercentStopUsesTheExcelCompatibleBaseCfvoShape()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.Scale("A1:A10", new ColorScale
        {
            Values =
            [
                ConditionalValue.Minimum,
                new ConditionalValue(ConditionalValueType.Percent, "HALF()"),
                ConditionalValue.Maximum,
            ],
            Colors =
            [
                ExcelColor.FromRgb(1, 2, 3),
                ExcelColor.FromRgb(4, 5, 6),
                ExcelColor.FromRgb(7, 8, 9),
            ],
        }));

        using MemoryStream stream = await SaveAsync(book);
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");

            WorksheetPart part = Assert.Single(document.WorkbookPart!.WorksheetParts);
            using var reader = new StreamReader(part.GetStream());
            string xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
            Assert.Contains("<cfvo type=\"percent\" val=\"HALF()\"/>", xml, StringComparison.Ordinal);
        }

        stream.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        ConditionalValue middle = Assert.IsType<ColorScale>(
            Assert.Single(reloaded.Sheets[0].ConditionalFormats).ColorScale).Values[1];

        Assert.Equal(ConditionalValueType.Percent, middle.Type);
        Assert.Equal("HALF()", middle.Value);
    }

    [Fact]
    public async Task DataBarLengthsAndThemeColorsRoundTripWithoutInvalidGte()
    {
        Workbook book = Sample();
        ExcelColor color = ExcelColor.FromTheme(5, -0.25);
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.Bars("A1:A10", new DataBar
        {
            Minimum = ConditionalValue.Number(0),
            Maximum = ConditionalValue.Number(1000),
            Color = color,
            MinimumLength = 15,
            MaximumLength = 80,
        }));

        using MemoryStream stream = await SaveAsync(book);
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");

            WorksheetPart part = Assert.Single(document.WorkbookPart!.WorksheetParts);
            using var reader = new StreamReader(part.GetStream());
            string xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);

            Assert.Contains("minLength=\"15\"", xml, StringComparison.Ordinal);
            Assert.Contains("maxLength=\"80\"", xml, StringComparison.Ordinal);
            Assert.DoesNotContain(" gte=", xml, StringComparison.Ordinal);
            Assert.Contains("<color theme=\"5\" tint=\"-0.25\"/>", xml, StringComparison.Ordinal);
        }

        stream.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        DataBar bar = Assert.IsType<DataBar>(Assert.Single(reloaded.Sheets[0].ConditionalFormats).DataBar);

        Assert.Equal(15, bar.MinimumLength);
        Assert.Equal(80, bar.MaximumLength);
        Assert.True(bar.Maximum.GreaterThanOrEqual);
        Assert.Equal(color, bar.Color);
    }

    [Fact]
    public async Task IconSetExclusiveThresholdRoundTrips()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.Icons("A1:A10", new IconSet
        {
            Name = "3TrafficLights1",
            Values =
            [
                ConditionalValue.Percent(0),
                ConditionalValue.Percent(33) with { GreaterThanOrEqual = false },
                ConditionalValue.Percent(67),
            ],
        }));

        using MemoryStream stream = await SaveAsync(book);
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");

            WorksheetPart part = Assert.Single(document.WorkbookPart!.WorksheetParts);
            using var reader = new StreamReader(part.GetStream());
            string xml = await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
            Assert.Contains(
                "<cfvo type=\"percent\" val=\"33\" gte=\"0\"/>",
                xml,
                StringComparison.Ordinal);
        }

        stream.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        IconSet icons = Assert.IsType<IconSet>(
            Assert.Single(reloaded.Sheets[0].ConditionalFormats).IconSet);

        Assert.True(icons.Values[0].GreaterThanOrEqual);
        Assert.False(icons.Values[1].GreaterThanOrEqual);
        Assert.True(icons.Values[2].GreaterThanOrEqual);
    }

    [Theory]
    [InlineData(-1, 90)]
    [InlineData(10, 101)]
    [InlineData(80, 20)]
    public void InvalidDataBarLengthsAreRejected(int minimum, int maximum)
    {
        Workbook book = Sample();
        ConditionalFormatRule rule = ConditionalFormatRule.Bars("A1:A10", new DataBar
        {
            MinimumLength = minimum,
            MaximumLength = maximum,
        });

        ArgumentException error = Assert.Throws<ArgumentException>(
            () => book.Sheets[0].ConditionalFormats.Add(rule));
        Assert.Contains("0 <= MinimumLength <= MaximumLength <= 100", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ClosedXmlSeesTheRulesAndTheirFormatting()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        IXLConditionalFormat format = Assert.Single(interop.Worksheet("Data").ConditionalFormats);
        Assert.Equal(XLConditionalFormatType.CellIs, format.ConditionalFormatType);
        Assert.Equal(XLCFOperator.GreaterThan, format.Operator);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFFFC7CE)), format.Style.Fill.BackgroundColor);
    }

    [Fact]
    public async Task StopIfTrueAndPrioritySurvive()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A10")],
            Type = ConditionalFormatType.CellIs,
            Operator = ConditionalOperator.LessThan,
            Formulas = ["300"],
            Style = DifferentialStyle.Good,
            Priority = 5,
            StopIfTrue = true,
        });

        ConditionalFormatRule rule = Assert.Single((await SaveAndReloadAsync(book)).Sheets[0].ConditionalFormats);

        Assert.Equal(5, rule.Priority);
        Assert.True(rule.StopIfTrue);
    }

    [Fact]
    public async Task AddingARuleKeepsTheOnesTheModelCannotDescribe()
    {
        // An extension-only rule is what Excel writes for the modern icon sets. It has to come
        // back untouched even though the sheet's rules were regenerated around it.
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Data");
        sheet["A1"] = 1;
        sheet.ConditionalFormats.AddFromFile(new ConditionalFormatRule
        {
            Ranges = [RangeRef.Parse("A1:A10")],
            RawXml = """
                <cfRule type="dataBar" priority="9" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>
                <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"/></extLst></cfRule>
                """,
        });

        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B1:B10", 1, DifferentialStyle.Bad));
        Worksheet loaded = (await SaveAndReloadAsync(book)).Sheets["Data"];

        Assert.Equal(2, loaded.ConditionalFormats.Count);
        Assert.True(loaded.ConditionalFormats[0].IsOpaque);
        Assert.Equal(ConditionalFormatType.CellIs, loaded.ConditionalFormats[1].Type);
    }

    [Fact]
    public async Task RulesReadFromAFileKeepTheirOriginalFormattingIndex()
    {
        Workbook book = Sample();
        book.Sheets[0].ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A10", 500, DifferentialStyle.Bad));
        Workbook reloaded = await SaveAndReloadAsync(book);

        // Adding a second rule regenerates the block; the first must keep pointing at the
        // differential format that came with the file rather than losing its colors.
        reloaded.Sheets[0].ConditionalFormats.Add(
            ConditionalFormatRule.CellLessThan("A1:A10", 100, DifferentialStyle.Good));

        using var stream = new MemoryStream();
        await reloaded.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        List<IXLConditionalFormat> formats = [.. interop.Worksheet("Data").ConditionalFormats];
        Assert.Equal(2, formats.Count);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFFFC7CE)), formats[0].Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.FromArgb(unchecked((int)0xFFC6EFCE)), formats[1].Style.Fill.BackgroundColor);
    }
}
