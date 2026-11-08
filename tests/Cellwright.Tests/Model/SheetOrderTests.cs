using System.IO.Compression;
using System.Text;
using Cellwright.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Reordering and duplicating sheets. Both are everyday operations in Excel and neither had an
/// equivalent here: the sheet list only knew how to grow at the end and shrink by name.
/// </summary>
public sealed class SheetOrderTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

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

    private static Workbook ThreeSheets()
    {
        Workbook workbook = Workbook.Create();
        foreach (string name in new[] { "First", "Second", "Third" })
            workbook.Sheets.Add(name)["A1"] = name;
        return workbook;
    }

    [Fact]
    public async Task MovingASheetChangesTheOrderInTheSavedFile()
    {
        Workbook workbook = ThreeSheets();
        workbook.Sheets.Move("Third", 0);

        Workbook reloaded = await SaveAndReloadAsync(workbook);
        Assert.Equal(["Third", "First", "Second"], reloaded.Sheets.Select(static sheet => sheet.Name));
    }

    [Fact]
    public void MovingASheetToWhereItAlreadyIsChangesNothing()
    {
        Workbook workbook = ThreeSheets();
        workbook.Sheets.Move("Second", 1);
        Assert.Equal(["First", "Second", "Third"], workbook.Sheets.Select(static sheet => sheet.Name));
    }

    [Fact]
    public void MovingRejectsAPositionOutsideTheWorkbookAndAForeignSheet()
    {
        Workbook workbook = ThreeSheets();
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.Sheets.Move("First", 3));
        Assert.Throws<ArgumentException>(() => workbook.Sheets.Move(Workbook.Create().Sheets.Add("Alien"), 0));
    }

    [Fact]
    public void ACopyLandsRightAfterItsOriginal()
    {
        Workbook workbook = ThreeSheets();
        workbook.Sheets.Copy("First", "First (2)");
        Assert.Equal(["First", "First (2)", "Second", "Third"], workbook.Sheets.Select(static sheet => sheet.Name));
    }

    [Fact]
    public async Task ACopyCarriesTheContentTheModelRepresents()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        source["A1"] = "Product";
        source["B1"] = 1234.5;
        source.Cell("B2").SetFormula("B1*2");
        source.Range("A1:B1").Style(static style => style.Bold());
        source.Column("A").SetWidth(24);
        source.Row(3).Hide();
        source.Merge(RangeRef.Parse("A5:B5"));
        source.FrozenRows = 1;
        source.ShowGridLines = false;
        source.ZoomScale = 85;
        source.AutoFilter("A1:B4");
        source.Comments.Add("A1", "check this", "Reviewer");
        source.Hyperlinks.Add(RangeRef.Parse("A2:A2"), "https://example.com", "docs");
        source.Validations.Add(DataValidation.WholeBetween("C1:C9", 1, 10));
        source.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("B1:B9", 100, DifferentialStyle.Good));
        source.PageSetup.Orientation = PageOrientation.Landscape;
        source.Protection.Protect();

        workbook.Sheets.Copy(source, "Report copy");
        Worksheet copy = (await SaveAndReloadAsync(workbook)).Sheets["Report copy"];

        Assert.Equal("Product", copy["A1"].GetText());
        Assert.Equal(1234.5, copy["B1"].GetNumber());
        Assert.Equal("B1*2", copy.Cell("B2").Formula);
        Assert.True(copy.Cell("A1").StyleValue.Font.Bold);
        Assert.Equal(24, copy.ConfiguredColumns[1].Width);
        Assert.True(copy.ConfiguredRows[3].Hidden);
        Assert.Equal(RangeRef.Parse("A5:B5"), Assert.Single(copy.MergedRanges));
        Assert.Equal(1, copy.FrozenRows);
        Assert.False(copy.ShowGridLines);
        Assert.Equal(85, copy.ZoomScale);
        Assert.Equal(RangeRef.Parse("A1:B4"), copy.AutoFilterRange);
        Assert.Equal("check this", copy.Comments[CellRef.Parse("A1")]!.Text);
        Assert.Equal("https://example.com", Assert.Single(copy.Hyperlinks).Target);
        Assert.Single(copy.Validations);
        Assert.Single(copy.ConditionalFormats);
        Assert.Equal(PageOrientation.Landscape, copy.PageSetup.Orientation);
        Assert.True(copy.Protection.IsEnabled);
    }

    [Fact]
    public void ACopyRetargetsExplicitSelfReferencesButLeavesExternalBooksAlone()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        source["A1"] = 1;
        source.Cell("A2").SetFormula(
            "Report!B1+'Report'!C1+[1]Report!D1+Other!E1+\"Report!F1\"");
        source.Validations.Add(DataValidation.Custom("B2", "Report!A1>0"));
        source.ConditionalFormats.Add(ConditionalFormatRule.Expression(
            "C2", "Report!A1>0", DifferentialStyle.Good));
        source.Hyperlinks.Add(Hyperlink.ToLocation(RangeRef.Parse("D2"), "Report!A1"));
        source.AddChart(
            ChartKind.Column,
            "F2:M17",
            [new ChartSeriesDefinition("Report!$A$1:$A$2", "[1]Report!$B$1:$B$2")]);

        SheetCopyResult result = workbook.Sheets.Copy(source, "Report copy", new SheetCopyOptions());
        Worksheet copy = result.Worksheet;

        Assert.False(result.HasWarnings);
        Assert.Equal(
            "'Report copy'!B1+'Report copy'!C1+[1]Report!D1+Other!E1+\"Report!F1\"",
            copy.Cell("A2").Formula);
        Assert.Equal("'Report copy'!A1>0", Assert.Single(copy.Validations).Formula1);
        Assert.Equal("'Report copy'!A1>0", Assert.Single(copy.ConditionalFormats).Formulas[0]);
        Assert.Equal("'Report copy'!A1", Assert.Single(copy.Hyperlinks).Location);

        ChartSeries series = Assert.Single(Assert.Single(copy.Charts).Series);
        Assert.Equal("'Report copy'!$A$1:$A$2", series.ValuesReference);
        Assert.Equal("[1]Report!$B$1:$B$2", series.CategoriesReference);
    }

    [Fact]
    public void SelfReferenceRetargetingCanBeDisabledExplicitly()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        source.Cell("A1").SetFormula("Report!B1");

        Worksheet copy = workbook.Sheets.Copy(
            source,
            "Report copy",
            new SheetCopyOptions { RetargetSelfReferences = false }).Worksheet;

        Assert.Equal("Report!B1", copy.Cell("A1").Formula);
    }

    [Fact]
    public async Task ACopyCarriesIndependentPicturesAndChartsThroughSave()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        source["A1"] = 10d;
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        source.Pictures.Add(png, "B2", "C3", "Logo");
        source.AddChart(ChartKind.Line, "E2:L17", "Report!$A$1:$A$2", title: "Trend");

        Worksheet copy = workbook.Sheets.Copy(source, "Report copy");

        Picture sourcePicture = Assert.Single(source.Pictures);
        Picture copiedPicture = Assert.Single(copy.Pictures);
        Assert.NotSame(sourcePicture, copiedPicture);
        Assert.NotSame(sourcePicture.Anchor, copiedPicture.Anchor);
        Assert.True(sourcePicture.Content.Span.SequenceEqual(copiedPicture.Content.Span));

        Chart sourceChart = Assert.Single(source.Charts);
        Chart copiedChart = Assert.Single(copy.Charts);
        Assert.NotSame(sourceChart, copiedChart);
        Assert.NotSame(sourceChart.Anchor, copiedChart.Anchor);
        Assert.Equal("'Report copy'!$A$1:$A$2", Assert.Single(copiedChart.Series).ValuesReference);

        copiedPicture.Anchor.From = CellRef.Parse("C4");
        copiedChart.Anchor.From = CellRef.Parse("F5");
        Assert.Equal(CellRef.Parse("B2"), sourcePicture.Anchor.From);
        Assert.Equal(CellRef.Parse("E2"), sourceChart.Anchor.From);

        Workbook reloaded = await SaveAndReloadAsync(workbook);
        Worksheet savedCopy = reloaded.Sheets["Report copy"];
        Assert.Equal("Logo", Assert.Single(savedCopy.Pictures).Name);
        Assert.Equal(CellRef.Parse("C4"), savedCopy.Pictures[0].Anchor.From);
        Assert.Equal("Trend", Assert.Single(savedCopy.Charts).Title);
        Assert.Equal(CellRef.Parse("F5"), savedCopy.Charts[0].Anchor.From);
        Assert.Equal("'Report copy'!$A$1:$A$2", savedCopy.Charts[0].Series[0].ValuesReference);
    }

    [Fact]
    public async Task ALoadedChartCopyOwnsItsRawPartAfterTheSourceIsRemoved()
    {
        Workbook created = Workbook.Create();
        Worksheet source = created.Sheets.Add("Report");
        source["A1"] = 10d;
        source["A2"] = 20d;
        source.AddChart(
            ChartKind.Column,
            "D2:K17",
            "Report!$A$1:$A$2",
            title: "Loaded chart",
            seriesName: "Amount");

        Workbook loaded = await SaveAndLoadWithExternalChartRelationshipAsync(created);
        Worksheet loadedSource = loaded.Sheets["Report"];
        Chart loadedChart = Assert.Single(loadedSource.Charts);
        Assert.NotNull(loadedChart.SourcePartPath);
        Assert.Single(loaded.Preserved!.RelationshipsOf(loadedChart.SourcePartPath!));

        SheetCopyResult result = loaded.Sheets.Copy(
            loadedSource,
            "Report copy",
            new SheetCopyOptions());
        Assert.False(result.HasWarnings);
        Assert.True(loaded.Sheets.Remove("Report"));

        using var saved = new MemoryStream();
        await loaded.SaveAsync(
            saved,
            new SaveOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        saved.Position = 0;
        using (var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true))
        {
            ZipArchiveEntry? relationshipPart = archive.GetEntry("xl/charts/_rels/chart2.xml.rels");
            Assert.NotNull(relationshipPart);
            using var reader = new StreamReader(relationshipPart.Open(), Encoding.UTF8);
            string relationships = reader.ReadToEnd();
            Assert.Contains("https://example.com/chart-image.png", relationships, StringComparison.Ordinal);
            Assert.Contains("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
        }

        saved.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(saved, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        saved.Position = 0;
        Workbook reloaded = await Workbook.LoadAsync(
            saved,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
        Worksheet copy = Assert.Single(reloaded.Sheets);
        Chart chart = Assert.Single(copy.Charts);
        ChartSeries series = Assert.Single(chart.Series);

        Assert.Equal("Report copy", copy.Name);
        Assert.Equal("Loaded chart", chart.Title);
        Assert.Equal("Amount", series.Name);
        Assert.Equal("'Report copy'!$A$1:$A$2", series.ValuesReference);
        Assert.Equal([10d, 20d], series.Values);
    }

    [Fact]
    public async Task LoadedOpaqueMarkupProducesAggregatedMachineReadableCopyDiagnostics()
    {
        Workbook workbook = await LoadWithOpaqueSheetMarkupAsync();
        Worksheet source = workbook.Sheets["Report"];
        var optionWarnings = new List<WorkbookWarning>();
        var workbookWarnings = new List<WorkbookWarning>();
        workbook.OnWarning = workbookWarnings.Add;

        SheetCopyResult result = workbook.Sheets.Copy(
            source,
            "Report copy",
            new SheetCopyOptions { OnWarning = optionWarnings.Add });

        WorkbookWarning warning = Assert.Single(result.Warnings);
        Assert.Equal("SheetCopyContentNotCopied", warning.Code);
        Assert.Contains("opaque worksheet markup", warning.Message, StringComparison.Ordinal);
        Assert.Equal(warning, Assert.Single(optionWarnings));
        Assert.Equal(warning, Assert.Single(workbookWarnings));
        Assert.Null(result.Worksheet.Markup);

        // The established overload remains source-compatible and surfaces the same limitation
        // through the workbook's edit-warning channel.
        Worksheet legacyCopy = workbook.Sheets.Copy(source, "Legacy copy");
        Assert.Equal("Legacy copy", legacyCopy.Name);
        Assert.Equal(2, workbookWarnings.Count);
        Assert.All(workbookWarnings, item => Assert.Equal("SheetCopyContentNotCopied", item.Code));

        Workbook reloaded = await SaveAndReloadAsync(workbook);
        Assert.Null(reloaded.Sheets["Report copy"].Markup?.SheetPropertyChildren);
        Assert.NotNull(reloaded.Sheets["Report"].Markup?.SheetPropertyChildren);
    }

    [Fact]
    public async Task ACopyCarriesAnIndependentThreadedConversation()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        Person nadia = workbook.People.Add("Nadia");
        Person tom = workbook.People.Add("Tom");
        ThreadedComment root = source.ThreadedComments.Add(CellRef.Parse("B2"), "Net or gross?", nadia);
        source.ThreadedComments.Reply(root, "Net.", tom);
        root.Resolved = true;

        // Normalize timestamps through the format before comparing a source part with its copy.
        Workbook loaded = await SaveAndReloadAsync(workbook);
        Worksheet loadedSource = loaded.Sheets["Report"];
        ThreadedComment sourceRoot = loadedSource.ThreadedComments[0];
        ThreadedComment sourceReply = loadedSource.ThreadedComments[1];

        Worksheet copy = loaded.Sheets.Copy(loadedSource, "Report copy");
        ThreadedComment copyRoot = copy.ThreadedComments[0];
        ThreadedComment copyReply = copy.ThreadedComments[1];

        Assert.NotSame(sourceRoot, copyRoot);
        Assert.NotSame(sourceReply, copyReply);
        Assert.NotEqual(sourceRoot.Id, copyRoot.Id);
        Assert.NotEqual(sourceReply.Id, copyReply.Id);
        Assert.Equal(copyRoot.Id, copyReply.ParentId);
        Assert.NotEqual(sourceRoot.Id, copyReply.ParentId);
        Assert.Equal(sourceRoot.Cell, copyRoot.Cell);
        Assert.Equal(sourceReply.Cell, copyReply.Cell);
        Assert.Equal(sourceRoot.Text, copyRoot.Text);
        Assert.Equal(sourceReply.Text, copyReply.Text);
        Assert.Equal(sourceRoot.Created, copyRoot.Created);
        Assert.Equal(sourceReply.Created, copyReply.Created);
        Assert.Equal(sourceRoot.PersonId, copyRoot.PersonId);
        Assert.Equal(sourceReply.PersonId, copyReply.PersonId);
        Assert.True(copyRoot.Resolved);
        Assert.False(copyReply.Resolved);

        copyRoot.Text = "Copied question";
        Assert.Equal("Net or gross?", sourceRoot.Text);

        Workbook reloaded = await SaveAndReloadAsync(loaded);
        Worksheet savedSource = reloaded.Sheets["Report"];
        Worksheet savedCopy = reloaded.Sheets["Report copy"];
        ThreadedComment savedCopyRoot = savedCopy.ThreadedComments[0];
        ThreadedComment savedCopyReply = savedCopy.ThreadedComments[1];

        Assert.Equal("Net or gross?", savedSource.ThreadedComments[0].Text);
        Assert.Equal("Copied question", savedCopyRoot.Text);
        Assert.Equal("Net.", savedCopyReply.Text);
        Assert.Equal(savedCopyRoot.Id, savedCopyReply.ParentId);
        Assert.True(savedCopyRoot.Resolved);
        Assert.Equal(4, reloaded.Sheets.SelectMany(static sheet => sheet.ThreadedComments)
            .Select(static comment => comment.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count());
    }

    /// <summary>
    /// Every mutable thing a sheet owns is copied by value. Sharing one would mean an insert on
    /// the copy silently moving a rule on the original.
    /// </summary>
    [Fact]
    public void EditingACopyLeavesTheOriginalAlone()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Report");
        source["A5"] = 1;
        source.Validations.Add(DataValidation.WholeBetween("A5:A9", 1, 10));
        source.ConditionalFormats.Add(ConditionalFormatRule.CellGreaterThan("A5:A9", 3, DifferentialStyle.Good));
        source.Hyperlinks.Add(RangeRef.Parse("A5:A5"), "https://example.com");
        source.Comments.Add("A5", "note");

        Worksheet copy = workbook.Sheets.Copy(source, "Report copy");
        copy.InsertRows(1, 3);
        copy["A1"] = "added";

        Assert.Equal(RangeRef.Parse("A5:A9"), Assert.Single(source.Validations).Ranges[0]);
        Assert.Equal(RangeRef.Parse("A5:A9"), Assert.Single(source.ConditionalFormats).Ranges[0]);
        Assert.Equal(RangeRef.Parse("A5:A5"), Assert.Single(source.Hyperlinks).Range);
        Assert.Equal(CellRef.Parse("A5"), Assert.Single(source.Comments).Cell);
        Assert.True(source["A1"].IsBlank);
        Assert.Equal(1, source["A5"].GetNumber());
        Assert.Equal(1, copy["A8"].GetNumber());
    }

    /// <summary>A table name is unique across the workbook, so the copy cannot keep the name.</summary>
    [Fact]
    public void ACopiedTableGetsANameOfItsOwn()
    {
        Workbook workbook = Workbook.Create();
        Worksheet source = workbook.Sheets.Add("Data");
        source["A1"] = "Product";
        source["B1"] = "Amount";
        source.Tables.Add("Sales", "A1:B4");

        Worksheet copy = workbook.Sheets.Copy(source, "Data copy");
        Assert.Equal("Sales", Assert.Single(source.Tables).Name);
        Assert.Equal("Sales2", Assert.Single(copy.Tables).Name);
    }

    [Fact]
    public void CopyingRejectsANameAlreadyInUse()
    {
        Workbook workbook = ThreeSheets();
        Assert.Throws<ArgumentException>(() => workbook.Sheets.Copy("First", "Second"));
    }

    private static async Task<Workbook> LoadWithOpaqueSheetMarkupAsync()
    {
        Workbook workbook = Workbook.Create();
        workbook.Sheets.Add("Report")["A1"] = "value";

        using var original = new MemoryStream();
        await workbook.SaveAsync(
            original,
            new SaveOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        original.Position = 0;
        var rewritten = new MemoryStream();
        using (var input = new ZipArchive(original, ZipArchiveMode.Read, leaveOpen: true))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in input.Entries)
            {
                ZipArchiveEntry destination = output.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using Stream target = destination.Open();
                using Stream source = entry.Open();
                if (entry.FullName != "xl/worksheets/sheet1.xml")
                {
                    source.CopyTo(target);
                    continue;
                }

                using var reader = new StreamReader(source, Encoding.UTF8);
                string xml = reader.ReadToEnd();
                int sheetData = xml.IndexOf("<sheetData", StringComparison.Ordinal);
                Assert.True(sheetData >= 0);
                string withOutline = xml.Insert(
                    sheetData,
                    "<sheetPr><outlinePr summaryBelow=\"0\"/></sheetPr>");
                target.Write(Encoding.UTF8.GetBytes(withOutline));
            }
        }

        rewritten.Position = 0;
        return await Workbook.LoadAsync(
            rewritten,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
    }

    private static async Task<Workbook> SaveAndLoadWithExternalChartRelationshipAsync(Workbook workbook)
    {
        using var original = new MemoryStream();
        await workbook.SaveAsync(
            original,
            new SaveOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);

        original.Position = 0;
        var rewritten = new MemoryStream();
        using (var input = new ZipArchive(original, ZipArchiveMode.Read, leaveOpen: true))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in input.Entries)
            {
                ZipArchiveEntry destination = output.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using Stream target = destination.Open();
                using Stream source = entry.Open();
                source.CopyTo(target);
            }

            ZipArchiveEntry relationships = output.CreateEntry(
                "xl/charts/_rels/chart1.xml.rels",
                CompressionLevel.Fastest);
            using Stream relationshipStream = relationships.Open();
            relationshipStream.Write(Encoding.UTF8.GetBytes($$"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId99" Type="{{XlsxSchema.RelImage}}"
                    Target="https://example.com/chart-image.png" TargetMode="External"/>
                </Relationships>
                """));
        }

        rewritten.Position = 0;
        return await Workbook.LoadAsync(
            rewritten,
            new LoadOptions { LeaveOpen = true },
            TestContext.Current.CancellationToken);
    }
}
