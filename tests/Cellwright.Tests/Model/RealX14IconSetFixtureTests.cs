using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Integration coverage for x14 icon sets emitted by real spreadsheet producers. The copied
/// fixtures are intentionally small: LibreOffice supplies all three post-2007 built-in sets,
/// while the NPOI corpus file exercises one of them beside base rules and x14 companions.
/// </summary>
public sealed class RealX14IconSetFixtureTests
{
    [Fact]
    public async Task LibreOfficeProjectsItsThreeBuiltInIconSetsExactly()
    {
        Workbook book = await LoadAsync("complex_icon_set.xlsx");
        Worksheet sheet = book.Sheets["new icon sets"];
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            sheet.Markup!.ExtensionConditionalFormats);

        Assert.False(projection.HasUnsupportedRules);
        Assert.Equal([1, 2, 3], sheet.Markup.ExtensionConditionalFormatPriorities!.Order());
        Assert.Empty(sheet.ConditionalFormats);

        Assert.Collection(
            projection.Rules,
            rule => AssertRule(rule, 3, "B2:B5", "3Triangles", ["0", "33", "67"]),
            rule => AssertRule(rule, 2, "D2:D5", "3Stars", ["0", "33", "67"]),
            rule => AssertRule(rule, 1, "F2:F6", "5Boxes", ["0", "20", "40", "60", "80"]));

        var evaluator = new ConditionalFormatEvaluator(sheet);
        AssertBands(evaluator, "B", 2, [0, 1, 1, 2], "3Triangles");
        AssertBands(evaluator, "D", 2, [0, 1, 1, 2], "3Stars");
        AssertBands(evaluator, "F", 2, [0, 1, 2, 3, 4], "5Boxes");
        Assert.Empty(evaluator.Unsupported);
    }

    [Fact]
    public async Task NpoiFixtureProjectsTheFourthBuiltInRuleBesideBaseAndOpaqueRules()
    {
        Workbook book = await LoadAsync("NewStyleConditionalFormattings.xlsx");
        Worksheet sheet = book.Sheets["CF"];
        var projection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            sheet.Markup!.ExtensionConditionalFormats);

        // The x14 data-bar companion has no priority and therefore does not reserve one.
        Assert.Equal([1, 6], sheet.Markup.ExtensionConditionalFormatPriorities!.Order());
        Assert.True(projection.HasUnsupportedRules);
        ConditionalFormatRule stars = Assert.Single(projection.Rules);
        AssertRule(stars, 6, "Q2:Q17", "3Stars", ["0", "33", "67"]);

        var evaluator = new ConditionalFormatEvaluator(sheet);
        AssertBands(
            evaluator,
            "Q",
            2,
            [0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 1, 0, 0, 0, 0, 0],
            "3Stars");
        Assert.Contains("conditional-format-opaque", evaluator.Unsupported);
    }

    [Fact]
    public async Task AllThreeCustomIconSetsRemainOpaque()
    {
        Assert.Equal(
            [
                "sheet2.xml|2|B2:B6|5Quarters|percent:0,percent:20,percent:40,percent:60,percent:80|" +
                    "3ArrowsGray:0,NoIcons:0,3Arrows:1,NoIcons:0,3Arrows:2",
                "sheet2.xml|1|D2:D4|3ArrowsGray|percent:0,percent:33,percent:67|" +
                    "4RedToBlack:3,3TrafficLights1:1,3Arrows:2",
                "sheet1.xml|1|U2:U17|(default)|percent:0,num:0,num:30|" +
                    "3Signs:0,3Flags:1,3Symbols2:2",
            ],
            ReadCustomIconSets());

        Workbook libreOffice = await LoadAsync("complex_icon_set.xlsx");
        Worksheet customSheet = libreOffice.Sheets["custom"];
        var customProjection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            customSheet.Markup!.ExtensionConditionalFormats);
        Assert.Empty(customProjection.Rules);
        Assert.True(customProjection.HasUnsupportedRules);
        Assert.Contains(
            "conditional-format-opaque",
            new ConditionalFormatEvaluator(customSheet).Unsupported);

        Workbook npoi = await LoadAsync("NewStyleConditionalFormattings.xlsx");
        var npoiProjection = Assert.IsType<global::Cellwright.IO.ExtensionConditionalFormatProjection>(
            npoi.Sheets["CF"].Markup!.ExtensionConditionalFormats);
        Assert.DoesNotContain(
            npoiProjection.Rules.SelectMany(static rule => rule.Ranges),
            range => range == RangeRef.Parse("U2:U17"));
    }

    private static void AssertRule(
        ConditionalFormatRule rule,
        int priority,
        string range,
        string name,
        string[] thresholds)
    {
        Assert.Equal(priority, rule.Priority);
        Assert.Equal([RangeRef.Parse(range)], rule.Ranges);
        Assert.Equal(ConditionalFormatType.IconSet, rule.Type);

        IconSet icons = Assert.IsType<IconSet>(rule.IconSet);
        Assert.Equal(name, icons.Name);
        Assert.True(icons.ShowValue);
        Assert.False(icons.Reverse);
        Assert.Equal(
            thresholds.Select(static threshold => new ConditionalValue(ConditionalValueType.Percent, threshold)),
            icons.Values);
        Assert.All(icons.Values, static threshold => Assert.True(threshold.GreaterThanOrEqual));
    }

    private static void AssertBands(
        ConditionalFormatEvaluator evaluator,
        string column,
        int firstRow,
        int[] expected,
        string setName)
    {
        for (int offset = 0; offset < expected.Length; offset++)
        {
            ConditionalIcon icon = Assert.IsType<ConditionalIcon>(
                evaluator.Of(CellRef.Parse($"{column}{firstRow + offset}")).Icon);
            Assert.Equal(setName, icon.SetName);
            Assert.Equal(expected[offset], icon.Index);
            Assert.True(icon.ShowValue);
        }
    }

    private static async Task<Workbook> LoadAsync(string fixtureName) =>
        await Workbook.LoadAsync(
            FixturePath(fixtureName),
            cancellationToken: TestContext.Current.CancellationToken);

    private static string FixturePath(string fixtureName)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string besideProject = Path.Combine(
                directory.FullName,
                "Fixtures",
                "X14",
                fixtureName);
            if (File.Exists(besideProject))
                return besideProject;

            string beneathRepository = Path.Combine(
                directory.FullName,
                "tests",
                "Cellwright.Tests",
                "Fixtures",
                "X14",
                fixtureName);
            if (File.Exists(beneathRepository))
                return beneathRepository;

            directory = directory.Parent;
        }

        throw new FileNotFoundException($"The x14 integration fixture '{fixtureName}' was not found.");
    }

    private static string[] ReadCustomIconSets() =>
    [
        .. ReadCustomIconSets("complex_icon_set.xlsx", "xl/worksheets/sheet2.xml"),
        .. ReadCustomIconSets("NewStyleConditionalFormattings.xlsx", "xl/worksheets/sheet1.xml"),
    ];

    private static IEnumerable<string> ReadCustomIconSets(string fixtureName, string entryName)
    {
        XNamespace x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        XNamespace xm = "http://schemas.microsoft.com/office/excel/2006/main";

        using ZipArchive package = ZipFile.OpenRead(FixturePath(fixtureName));
        ZipArchiveEntry entry = package.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing package part '{entryName}'.");
        using Stream content = entry.Open();
        XDocument sheet = XDocument.Load(content, System.Xml.Linq.LoadOptions.PreserveWhitespace);

        foreach (XElement block in sheet.Descendants(x14 + "conditionalFormatting"))
        {
            XElement? rule = block.Element(x14 + "cfRule");
            XElement? icons = rule?.Element(x14 + "iconSet");
            if ((string?)icons?.Attribute("custom") != "1")
                continue;

            string thresholds = string.Join(",", icons!.Elements(x14 + "cfvo").Select(value =>
                $"{(string?)value.Attribute("type")}:{(string?)value.Element(xm + "f")}"));
            string customIcons = string.Join(",", icons.Elements(x14 + "cfIcon").Select(icon =>
                $"{(string?)icon.Attribute("iconSet")}:{(string?)icon.Attribute("iconId")}"));

            yield return string.Join(
                "|",
                Path.GetFileName(entryName),
                (string?)rule!.Attribute("priority"),
                (string?)block.Element(xm + "sqref"),
                (string?)icons.Attribute("iconSet") ?? "(default)",
                thresholds,
                customIcons);
        }
    }
}
