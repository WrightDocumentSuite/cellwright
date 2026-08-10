using Inkwright;
using Inkwright.Annotations;
using Inkwright.Cos;
using Inkwright.Layout;
using Inkwright.Tagging;
using Xunit;

namespace Cellwright.Pdf.Tests;

public sealed class CommentTests
{
    private static readonly DateTimeOffset RootTime =
        new(2026, 8, 8, 10, 15, 0, TimeSpan.Zero);

    [Fact]
    public void ClassicNoteIsOptInInteractiveAndCarriesNoInventedTime()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Anchor";
        sheet.Comments.Add("A1", "Check the total.", "Ada");
        sheet.Hyperlinks.Add(Hyperlink.ToUrl(RangeRef.Parse("A1"), "https://example.org/anchor"));

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        PdfPage page = rendered.Document.Pages[0];
        PdfTextAnnotation note = Assert.IsType<PdfTextAnnotation>(Assert.Single(page.Annotations.Comments));
        PdfRectangle cell = Assert.Single(page.Annotations.OfType<PdfLinkAnnotation>()).Rectangle;
        Assert.Equal("Check the total.", note.Contents);
        Assert.Equal("Ada", note.Author);
        Assert.Null(note.CreatedAt);
        Assert.Null(note.ModifiedAt);
        Assert.Equal($"cellwright-note-{sheet.Id}-1-1", note.Name);
        Assert.Equal("Excel note", note.Subject);
        Assert.Single(rendered.Document.Pages[0].Annotations.OfType<PdfPopupAnnotation>());
        Assert.DoesNotContain(rendered.Diagnostics, warning => warning.Subject is "comment" or "threaded-comments");

        (double textX, double textY) = Assert.IsType<(double X, double Y)>(rendered.Find("Anchor"));
        Assert.True(note.Rectangle.Right > textX);
        Assert.True(note.Rectangle.Top > textY);
        PdfRectangle crop = rendered.Document.Pages[0].CropBox;
        Assert.True(crop.Contains(note.Rectangle.BottomLeft));
        Assert.True(crop.Contains(note.Rectangle.TopRight));
        Assert.Equal(14, note.Rectangle.Width, 6);
        Assert.Equal(14, note.Rectangle.Height, 6);
        Assert.Equal(cell.Right, note.Rectangle.Center.X, 6);
        Assert.Equal(cell.Top, note.Rectangle.Center.Y, 6);
        Assert.True(note.Rectangle.Left < cell.Right && note.Rectangle.Right > cell.Right);
    }

    [Fact]
    public void ThreadedRepliesAndResolvedInformationSurviveSaveAndReload()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Anchor";

        Person ada = AddPerson(book, "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", "Ada");
        Person grace = AddPerson(book, "{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", "Grace");
        ThreadedComment root = AddMessage(
            sheet,
            "{10000000-0000-0000-0000-000000000001}",
            null,
            ada,
            "First review note.",
            RootTime,
            resolved: true);
        ThreadedComment reply = AddMessage(
            sheet,
            "{10000000-0000-0000-0000-000000000002}",
            root.Id,
            grace,
            "Agreed.",
            created: null,
            resolved: true);

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        PdfTextAnnotation note = Assert.IsType<PdfTextAnnotation>(
            Assert.Single(rendered.Document.Pages[0].Annotations.Comments));
        Assert.Equal("First review note.", note.Contents);
        Assert.Equal("Ada", note.Author);
        Assert.Equal(RootTime, note.CreatedAt);
        Assert.Equal(RootTime, note.ModifiedAt);
        Assert.Contains(root.Id, note.Name, StringComparison.Ordinal);
        Assert.Equal(14, note.Rectangle.Width, 6);
        Assert.Equal(14, note.Rectangle.Height, 6);

        PdfMarkupAnnotation actualReply = Assert.Single(note.Replies, candidate => candidate.Contents == "Agreed.");
        Assert.Equal(note.Id, actualReply.InReplyTo?.Id);
        Assert.Equal("Grace", actualReply.Author);
        Assert.Null(actualReply.CreatedAt);
        Assert.Null(actualReply.ModifiedAt);
        Assert.Contains(reply.Id, actualReply.Name, StringComparison.Ordinal);
        Assert.Empty(DrawnOperators(actualReply));

        PdfMarkupAnnotation rootResolution = Assert.Single(
            note.Replies,
            candidate => candidate.Contents?.StartsWith("Resolved in the source Excel", StringComparison.Ordinal) == true);
        AssertInformationalResolution(rootResolution, note);

        PdfMarkupAnnotation replyResolution = Assert.Single(actualReply.Replies);
        AssertInformationalResolution(replyResolution, actualReply);
        Assert.Equal(PdfReviewState.None, note.ReviewState);

        PdfExportWarning warning = Assert.Single(
            rendered.Diagnostics,
            candidate => candidate.Subject == "comment-resolved-state");
        Assert.Contains("2 resolved", warning.Message, StringComparison.Ordinal);
        Assert.Contains("/State /T", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ClassicAndThreadedCommentsCanShareOneCell()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Anchor";
        sheet.Comments.Add("A1", "Classic note", "Ada");
        Person person = AddPerson(book, "{ABABABAB-ABAB-ABAB-ABAB-ABABABABABAB}", "Grace");
        AddMessage(sheet, "thread-at-same-cell", null, person, "Threaded note", RootTime);

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        PdfMarkupAnnotation[] roots = [.. rendered.Document.Pages[0].Annotations.Comments];
        Assert.Equal(2, roots.Length);
        Assert.Contains(roots, note => note.Contents == "Classic note" && note.Author == "Ada");
        Assert.Contains(roots, note => note.Contents == "Threaded note" && note.Author == "Grace");
        Assert.Equal(2, roots.Select(note => note.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.NotEqual(roots[0].Rectangle, roots[1].Rectangle);
        Assert.Equal(2, rendered.Document.Pages[0].Annotations.OfType<PdfPopupAnnotation>().Count());
    }

    [Fact]
    public void UnsupportedCommentMetadataAndMissingAuthorsAreAggregatedWithoutXmlLeakage()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Anchor";
        book.People.ExtensionListXml = "<person-list secret='one'/>";
        var known = new Person("{ACACACAC-ACAC-ACAC-ACAC-ACACACACACAC}", "Ada")
        {
            ExtensionListXml = "<person secret='two'/>",
        };
        book.People.AddFromFile(known);
        sheet.ThreadedComments.ExtensionListXml = "<collection secret='three'/>";
        AddMessage(
            sheet,
            "opaque-message",
            null,
            known,
            "Mentions Ada",
            RootTime,
            mentionsXml: "<mentions secret='four'/>",
            extensionListXml: "<message secret='five'/>");
        var unknown = new Person("{ADADADAD-ADAD-ADAD-ADAD-ADADADADADAD}", "Not in workbook");
        AddMessage(sheet, "unknown-author", null, unknown, "Anonymous in PDF", null);

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        PdfExportWarning metadata = Assert.Single(
            rendered.Diagnostics,
            warning => warning.Subject == "comment-metadata");
        Assert.Contains("1 mention block", metadata.Message, StringComparison.Ordinal);
        Assert.Contains("1 message extension", metadata.Message, StringComparison.Ordinal);
        Assert.Contains("1 collection extension", metadata.Message, StringComparison.Ordinal);
        Assert.Contains("1 person extension", metadata.Message, StringComparison.Ordinal);
        Assert.Contains("1 person-list extension", metadata.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", metadata.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "comment-authors");
        PdfMarkupAnnotation anonymous = Assert.Single(
            rendered.Document.Pages[0].Annotations.Comments,
            note => note.Contents == "Anonymous in PDF");
        Assert.Null(anonymous.Author);
    }

    [Fact]
    public void CommentsDoNotChangePaginationAndRepeatedTitlesDoNotDuplicateThem()
    {
        Workbook book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");

        for (int row = 1; row <= 8; row++)
            first.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"Row {row}"));

        first.SetRepeatRows(1, 1);
        first.PageSetup.AddRowBreak(5);
        first.Comments.Add("A1", "Repeated title note");
        first.Comments.Add("A6", "Second page note");
        second["A1"] = "Second sheet";
        second.Comments.Add("A1", "Second sheet note");

        using Rendered printed = Rendered.Of(book);
        using Rendered review = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        Assert.Equal(printed.PageCount, review.PageCount);
        Assert.Equal(3, review.PageCount);
        Assert.Single(review.Document.Pages[0].Annotations.Comments);
        Assert.Single(review.Document.Pages[1].Annotations.Comments);
        Assert.Single(review.Document.Pages[2].Annotations.Comments);
        Assert.Equal(
            3,
            review.Document.Pages.SelectMany(page => page.Annotations.Comments).Select(note => note.Name).Distinct().Count());
    }

    [Fact]
    public void HiddenAndNonPrintedAnchorsAreDiagnosedSeparately()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Visible";
        sheet["A2"] = "Hidden";
        sheet.Row(2).Hidden = true;
        sheet.Comments.Add("A2", "Cannot be seen");
        sheet.Comments.Add("B1", "Outside the used print rectangle");

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        Assert.Empty(rendered.Document.Pages[0].Annotations.Comments);
        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "comment-hidden");
        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "comment-anchors");
    }

    [Fact]
    public void DanglingCyclesDuplicatesAndCrossCellRepliesAreSafeAndDiagnosed()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Malformed");
        sheet["A1"] = "A";
        sheet["A2"] = "B";
        Person person = AddPerson(book, "{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}", "Reviewer");

        AddMessage(sheet, "dangling", "missing", person, "Dangling", null);
        AddMessage(sheet, "cycle-a", "cycle-b", person, "Cycle A", null);
        AddMessage(sheet, "cycle-b", "cycle-a", person, "Cycle B", null);
        AddMessage(sheet, "duplicate", null, person, "Duplicate one", null);
        AddMessage(sheet, "duplicate", null, person, "Duplicate two", null);
        AddMessage(sheet, "cross-root", null, person, "Cross root", null);
        AddMessage(sheet, "cross-reply", "cross-root", person, "Cross reply", null, cell: "A2");

        using Rendered rendered = Rendered.Of(book, new PdfExportOptions { IncludeComments = true });

        Assert.Equal(7, rendered.Document.Pages[0].Annotations.OfType<PdfMarkupAnnotation>().Count());
        Assert.Contains(rendered.Diagnostics, warning => warning.Subject == "comment-thread-dangling");
        PdfExportWarning malformed = Assert.Single(
            rendered.Diagnostics,
            warning => warning.Subject == "comment-thread-malformed");
        Assert.Contains("duplicate", malformed.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cycle", malformed.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cross-cell", malformed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TenThousandDeepMessagesRenderWithoutRecursiveStackGrowth()
    {
        const int Count = 10_000;
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Deep");
        sheet["A1"] = "Anchor";
        Person person = AddPerson(book, "{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}", "Reviewer");

        string? parent = null;
        for (int i = 0; i < Count; i++)
        {
            string id = $"deep-{i}";
            AddMessage(sheet, id, parent, person, $"Message {i}", null);
            parent = id;
        }

        PdfExportResult result = PdfExporter.Render(book, new PdfExportOptions { IncludeComments = true });
        using PdfDocument document = result.Document;

        Assert.Equal(Count, document.Pages[0].Annotations.OfType<PdfMarkupAnnotation>().Count());
    }

    [Fact]
    public void TaggedCommentsHaveOrderedAnnotObjrAndAValidParentTree()
    {
        Workbook book = Workbook.Create();
        book.Properties.Title = "Accessible review";
        Worksheet sheet = book.Sheets.Add("Review");
        sheet["A1"] = "Anchor";
        Person person = AddPerson(book, "{EEEEEEEE-EEEE-EEEE-EEEE-EEEEEEEEEEEE}", "Ada");
        ThreadedComment root = AddMessage(sheet, "tag-root", null, person, "Root", RootTime);
        AddMessage(
            sheet,
            "tag-reply",
            root.Id,
            person,
            "Reply",
            RootTime.AddMinutes(1),
            resolved: true);

        var options = new PdfExportOptions
        {
            IncludeComments = true,
            Tagged = true,
            Language = "en-GB",
        };

        using Rendered rendered = Rendered.Of(
            book,
            options,
            document => PdfUaProfile.Declare(document, PdfUaConformance.Ua1, "Accessible review"));

        PdfPage page = rendered.Document.Pages[0];
        Assert.Equal("S", page.Dictionary.GetName(PdfName.Get("Tabs"))?.Value);

        PdfStructureNode documentNode = Assert.Single(rendered.Document.Structure);
        PdfStructureNode table = Assert.Single(documentNode.Children);
        PdfStructureNode row = Assert.Single(table.Children);
        PdfStructureNode cell = Assert.Single(row.Children);
        Assert.Equal("TD", cell.Tag);
        Assert.Collection(
            cell.Kids,
            first => Assert.IsType<PdfStructureContentKid>(first),
            second => AssertAnnot(second, rendered.Document),
            third => AssertAnnot(third, rendered.Document),
            fourth => AssertAnnot(fourth, rendered.Document));

        PdfMarkupAnnotation[] messages = [.. page.Annotations.OfType<PdfMarkupAnnotation>()];
        Assert.Equal(3, messages.Length);
        PdfMarkupAnnotation taggedReply = Assert.Single(messages, message => message.Contents == "Reply");
        PdfMarkupAnnotation taggedResolution = Assert.Single(taggedReply.Replies);
        AssertInformationalResolution(taggedResolution, taggedReply);
        Assert.All(messages, annotation =>
        {
            int key = (int)annotation.Dictionary.GetInteger(PdfName.Get("StructParent"), -1);
            Assert.True(key >= 0);
            PdfStructureNode owner = Assert.Single(
                cell.Children,
                child => child.Objects.Any(reference => reference.Object == annotation.Id));
            Assert.Equal(owner.Id, ParentTreeValue(rendered.Document, key).AsReference);
        });

        Assert.Empty(
            PdfUaProfile.Violations(PdfUaProfile.Validate(rendered.Document, PdfUaConformance.Ua1)));
    }

    [Fact]
    public void TaggedRepeatedTitleInterleavesOneAnnotBetweenPageContentAndMultiPageMcr()
    {
        Workbook book = Workbook.Create();
        book.Properties.Title = "Repeated review";
        Worksheet sheet = book.Sheets.Add("Review");
        for (int row = 1; row <= 8; row++)
            sheet.Cell(new CellRef(row, 1)).SetValue(CellValue.Text($"Row {row}"));

        sheet.SetRepeatRows(1, 1);
        sheet.PageSetup.AddRowBreak(5);
        sheet.Comments.Add("A1", "One note on the repeated title", "Ada");

        var options = new PdfExportOptions
        {
            IncludeComments = true,
            Tagged = true,
            Language = "en-GB",
        };
        using Rendered rendered = Rendered.Of(
            book,
            options,
            document => PdfUaProfile.Declare(document, PdfUaConformance.Ua1, "Repeated review"));

        Assert.Equal(2, rendered.PageCount);
        PdfPage firstPage = rendered.Document.Pages[0];
        PdfPage secondPage = rendered.Document.Pages[1];
        PdfTextAnnotation note = Assert.IsType<PdfTextAnnotation>(Assert.Single(firstPage.Annotations.Comments));
        Assert.Empty(secondPage.Annotations.Comments);

        PdfStructureNode table = Assert.Single(Assert.Single(rendered.Document.Structure).Children);
        PdfStructureNode titleRow = table.Children[0];
        PdfStructureNode cell = Assert.Single(titleRow.Children);
        Assert.Collection(
            cell.Kids,
            first => Assert.Equal(firstPage.Id, Assert.IsType<PdfStructureContentKid>(first).Reference.Page),
            second => AssertAnnot(second, rendered.Document),
            third => Assert.Equal(secondPage.Id, Assert.IsType<PdfStructureContentKid>(third).Reference.Page));

        PdfArray rawKids = Assert.IsType<PdfArray>(cell.Dictionary.Get(PdfName.Get("K")).AsArray);
        Assert.Equal(3, rawKids.Count);
        Assert.Equal(0, rawKids.GetInteger(0, -1));
        Assert.True(rawKids[1].IsReference);
        PdfDictionary mcr = Assert.IsType<PdfDictionary>(rawKids[2].AsDictionary);
        Assert.Equal("MCR", mcr.GetName(PdfName.Get("Type"))?.Value);
        Assert.Equal(secondPage.Id, mcr[PdfName.Get("Pg")].AsReference);
        Assert.Equal(
            Assert.IsType<PdfStructureContentKid>(cell.Kids[2]).Reference.MarkedContentId,
            mcr.GetInteger(PdfName.Get("MCID"), -1));

        AssertMcidParent(rendered.Document, firstPage, cell.Id, cell.Content[0].MarkedContentId);
        AssertMcidParent(rendered.Document, secondPage, cell.Id, cell.Content[1].MarkedContentId);
        PdfStructureNode annotOwner = Assert.IsType<PdfStructureElementKid>(cell.Kids[1]).Element;
        int annotationKey = (int)note.Dictionary.GetInteger(PdfName.Get("StructParent"), -1);
        Assert.Equal(annotOwner.Id, ParentTreeValue(rendered.Document, annotationKey).AsReference);

        Assert.Empty(
            PdfUaProfile.Violations(PdfUaProfile.Validate(rendered.Document, PdfUaConformance.Ua1)));
    }

    private static Person AddPerson(Workbook book, string id, string name)
    {
        var person = new Person(id, name);
        book.People.AddFromFile(person);
        return person;
    }

    private static ThreadedComment AddMessage(
        Worksheet sheet,
        string id,
        string? parentId,
        Person person,
        string text,
        DateTimeOffset? created,
        bool resolved = false,
        string cell = "A1",
        string? mentionsXml = null,
        string? extensionListXml = null)
    {
        var message = new ThreadedComment
        {
            Cell = CellRef.Parse(cell),
            Text = text,
            PersonId = person.Id,
            Id = id,
            ParentId = parentId,
            Created = created,
            Resolved = resolved,
            MentionsXml = mentionsXml,
            ExtensionListXml = extensionListXml,
        };
        sheet.ThreadedComments.AddFromFile(message);
        return message;
    }

    private static void AssertAnnot(PdfStructureKid kid, PdfDocument document)
    {
        PdfStructureNode node = Assert.IsType<PdfStructureElementKid>(kid).Element;
        Assert.Equal("Annot", node.Tag);
        PdfObjectReference reference = Assert.Single(node.Objects);
        PdfDictionary objr = Assert.IsType<PdfDictionary>(node.Dictionary.Get(PdfName.Get("K")).AsDictionary);
        Assert.Equal("OBJR", objr.GetName(PdfName.Get("Type"))?.Value);
        Assert.Equal(reference.Object, objr[PdfName.Get("Obj")].AsReference);
        Assert.Contains(document.AllAnnotations, annotation => annotation.Id == reference.Object);
    }

    private static void AssertInformationalResolution(
        PdfMarkupAnnotation resolution,
        PdfMarkupAnnotation expectedParent)
    {
        Assert.Equal(expectedParent.Id, resolution.InReplyTo?.Id);
        Assert.Null(resolution.Author);
        Assert.Null(resolution.CreatedAt);
        Assert.Null(resolution.ModifiedAt);
        Assert.True(resolution.Dictionary.Get(PdfName.Get("State")).IsNull);
        Assert.True(resolution.Dictionary.Get(PdfName.Get("StateModel")).IsNull);
        Assert.Empty(DrawnOperators(resolution));
    }

    private static string DrawnOperators(PdfAnnotation annotation)
    {
        PdfDictionary? appearance = annotation.Dictionary.GetDictionary(PdfName.Get("AP"));
        if (appearance?.GetStream(PdfName.Get("N")) is not { } stream)
            return string.Empty;

        string content = System.Text.Encoding.Latin1.GetString(stream.DecodedSpan());
        return content.Replace("q", string.Empty, StringComparison.Ordinal)
                      .Replace("Q", string.Empty, StringComparison.Ordinal)
                      .Trim();
    }

    private static PdfValue ParentTreeValue(PdfDocument document, int key)
    {
        PdfDictionary root = document.Catalog.GetDictionary(PdfName.Get("StructTreeRoot"))!;
        PdfDictionary tree = root.GetDictionary(PdfName.Get("ParentTree"))!;
        PdfArray numbers = tree.GetArray(PdfName.Get("Nums"))!;

        for (int i = 0; i + 1 < numbers.Count; i += 2)
        {
            if (numbers.GetInteger(i, -1) == key)
                return numbers[i + 1];
        }

        return PdfValue.Null;
    }

    private static void AssertMcidParent(
        PdfDocument document,
        PdfPage page,
        PdfObjectId expectedOwner,
        int mcid)
    {
        int pageKey = (int)page.Dictionary.GetInteger(PdfName.Get("StructParents"), -1);
        PdfArray parents = Assert.IsType<PdfArray>(document.Resolve(ParentTreeValue(document, pageKey)).AsArray);
        Assert.True(mcid >= 0 && mcid < parents.Count);
        Assert.Equal(expectedOwner, parents[mcid].AsReference);
    }

}
