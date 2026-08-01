using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Styles;

/// <summary>
/// What §18.8 calls the styles part must survive a round trip even where the model view of it
/// is partial: gradient fills, named cell styles and the differential formats of loaded rules.
/// </summary>
public sealed class StylePreservationTests
{
    [Fact]
    public async Task GradientFill_RoundTripsAndValidates()
    {
        var gradient = new GradientFill
        {
            Degree = 90,
            Stops =
            [
                new GradientStop(0, ExcelColor.FromArgb(0xFF92D050)),
                new GradientStop(1, ExcelColor.FromArgb(0xFF0070C0)),
            ],
        };

        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("G");
        sheet["A1"] = "gradient";
        sheet.Cell("A1").Style(s => s with { Fill = FillStyle.Gradiented(gradient) });

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            var errors = new OpenXmlValidator().Validate(document, TestContext.Current.CancellationToken).ToList();
            Assert.Empty(errors);
        }

        stream.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(stream, cancellationToken: TestContext.Current.CancellationToken);
        FillStyle fill = reloaded.Sheets[0].Cell("A1").StyleValue.Fill;
        Assert.Equal(gradient, fill.Gradient);
        Assert.False(fill.IsEmpty);
    }

    [Fact]
    public async Task GradientFill_PathVariant_KeepsConvergenceInsets()
    {
        var gradient = new GradientFill
        {
            Type = GradientFillType.Path,
            Left = 0.3,
            Right = 0.7,
            Top = 0.25,
            Bottom = 0.75,
            Stops =
            [
                new GradientStop(0, ExcelColor.FromArgb(0xFFFFFFFF)),
                new GradientStop(1, ExcelColor.FromTheme(4, -0.25)),
            ],
        };

        var book = Workbook.Create();
        book.Sheets.Add("G").Cell("B2").SetValue(1).Style(s => s with { Fill = new FillStyle { Gradient = gradient } });

        Workbook reloaded = await RoundTripAsync(book);
        Assert.Equal(gradient, reloaded.Sheets[0].Cell("B2").StyleValue.Fill.Gradient);
    }

    [Fact]
    public async Task NamedCellStyles_SurviveRoundTrip()
    {
        using MemoryStream original = await BuildBookWithNamedStyleAsync();
        Workbook book = await Workbook.LoadAsync(original, cancellationToken: TestContext.Current.CancellationToken);

        using var saved = new MemoryStream();
        await book.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string styles = ReadEntry(saved, "xl/styles.xml");
        Assert.Contains("cellStyleXfs count=\"2\"", styles, StringComparison.Ordinal);
        Assert.Contains("name=\"Accent\"", styles, StringComparison.Ordinal);

        // The styled cell keeps pointing at its named style, not at Normal.
        Assert.Contains("xfId=\"1\"", styles, StringComparison.Ordinal);

        saved.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(saved, isEditable: false);
        var errors = new OpenXmlValidator().Validate(document, TestContext.Current.CancellationToken).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public async Task ConditionalRuleFromFile_ExposesItsDifferentialStyle()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("CF");
        sheet["A1"] = 5;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A9", 3, DifferentialStyle.Bad));

        Workbook reloaded = await RoundTripAsync(book);
        ConditionalFormatRule rule = reloaded.Sheets[0].ConditionalFormats[0];

        Assert.NotNull(rule.Style);
        Assert.Equal(DifferentialStyle.Bad.FontColor, rule.Style!.FontColor);
        Assert.Equal(DifferentialStyle.Bad.FillColor, rule.Style.FillColor);
    }

    [Fact]
    public async Task UntouchedRuleStyle_KeepsItsFormatIndex()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("CF");
        sheet["A1"] = 5;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A9", 3, DifferentialStyle.Bad));

        Workbook reloaded = await RoundTripAsync(book);

        using var saved = new MemoryStream();
        await reloaded.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        string styles = ReadEntry(saved, "xl/styles.xml");

        // The parsed style compares equal to what the file already had, so no second dxf
        // is minted and the rule keeps its original index.
        Assert.Contains("<dxfs count=\"1\">", styles, StringComparison.Ordinal);
    }

    [Fact]
    public async Task EditedRuleStyle_GetsItsOwnFormat()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("CF");
        sheet["A1"] = 5;
        sheet.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A1:A9", 3, DifferentialStyle.Bad));

        Workbook reloaded = await RoundTripAsync(book);
        Worksheet loadedSheet = reloaded.Sheets[0];
        loadedSheet.ConditionalFormats[0].Style = DifferentialStyle.Good;

        using var saved = new MemoryStream();
        await reloaded.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        string styles = ReadEntry(saved, "xl/styles.xml");
        Assert.Contains("<dxfs count=\"2\">", styles, StringComparison.Ordinal);

        saved.Position = 0;
        Workbook again = await Workbook.LoadAsync(saved, cancellationToken: TestContext.Current.CancellationToken);
        Assert.Equal(DifferentialStyle.Good.FillColor, again.Sheets[0].ConditionalFormats[0].Style!.FillColor);
    }

    /// <summary>Builds a package whose styles part carries a named style used by cell A1.</summary>
    private static async Task<MemoryStream> BuildBookWithNamedStyleAsync()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = "named";
        sheet.Cell("A1").Style(s => s with { Font = s.Font with { Bold = true } });

        var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        string styles = ReadEntry(stream, "xl/styles.xml");
        styles = styles
            .Replace(
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
                "<cellStyleXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
                StringComparison.Ordinal)
            .Replace(
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
                "<cellStyles count=\"2\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" +
                "<cellStyle name=\"Accent\" xfId=\"1\"/></cellStyles>",
                StringComparison.Ordinal);

        // Attach the bold cell format to the new master.
        int boldXf = styles.IndexOf("applyFont=\"1\"", StringComparison.Ordinal);
        Assert.True(boldXf > 0);
        styles = styles[..boldXf].Replace("xfId=\"0\" applyFont", "xfId=\"1\" applyFont", StringComparison.Ordinal)
            + styles[boldXf..];

        return RewriteEntry(stream, "xl/styles.xml", styles);
    }

    private static async Task<Workbook> RoundTripAsync(Workbook book)
    {
        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;
        return await Workbook.LoadAsync(stream, cancellationToken: TestContext.Current.CancellationToken);
    }

    private static string ReadEntry(MemoryStream package, string entryName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(archive.GetEntry(entryName)!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>Copies the package with one entry's content replaced.</summary>
    private static MemoryStream RewriteEntry(MemoryStream package, string entryName, string content)
    {
        var rewritten = new MemoryStream();
        package.Position = 0;
        using (var source = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                using Stream output = target.CreateEntry(entry.FullName).Open();
                if (entry.FullName == entryName)
                {
                    output.Write(Encoding.UTF8.GetBytes(content));
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
}
