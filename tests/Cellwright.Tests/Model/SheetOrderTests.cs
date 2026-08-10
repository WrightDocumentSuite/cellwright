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
}
