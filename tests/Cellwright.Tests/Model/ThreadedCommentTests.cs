using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Cellwright.IO;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Threaded comments: the conversation attached to a cell, the people who wrote it, and the
/// legacy note that stands beside it for readers that predate threads.
/// </summary>
public sealed class ThreadedCommentTests
{
    private const string CanonicalPersonsProbe = "<probe xmlns=\"urn:cellwright:test\">keep</probe>";

    private static async Task<(Workbook Book, byte[] Bytes)> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        byte[] bytes = stream.ToArray();
        stream.Position = 0;
        Workbook read = await Workbook.LoadAsync(
            stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        return (read, bytes);
    }

    private static string PartOf(byte[] package, string suffix)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (!entry.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                continue;

            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            return reader.ReadToEnd();
        }

        return string.Empty;
    }

    private static bool HasPart(byte[] package, string suffix) => PartOf(package, suffix).Length > 0;

    private static IReadOnlyList<string> XmlPartsUnder(byte[] package, string directory)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var parts = new List<string>();
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (!entry.FullName.Contains(directory, StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                continue;

            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            parts.Add(reader.ReadToEnd());
        }

        return parts;
    }

    private static string SheetRelatedPart(byte[] package, int sheetIndex, string relationshipSuffix)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        string sheetPath = $"/xl/worksheets/sheet{sheetIndex}.xml";
        ZipArchiveEntry relationships = archive.GetEntry(
            $"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels")!;
        using Stream stream = relationships.Open();
        XDocument document = XDocument.Load(stream);
        XElement relationship = Assert.Single(
            document.Root!.Elements(),
            element => element.Attribute("Type")?.Value.EndsWith(relationshipSuffix, StringComparison.Ordinal) == true);
        string target = relationship.Attribute("Target")!.Value;
        string partPath = OpcPath.Resolve(sheetPath, target).TrimStart('/');
        using var reader = new StreamReader(archive.GetEntry(partPath)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static byte[] RewritePart(byte[] package, string suffix, Func<string, string> rewrite)
    {
        using var sourceStream = new MemoryStream(package);
        using var source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
        using var result = new MemoryStream();
        using (var target = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            bool found = false;
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                ZipArchiveEntry copy = target.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using Stream output = copy.Open();
                using Stream input = entry.Open();
                if (!entry.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    input.CopyTo(output);
                    continue;
                }

                found = true;
                using var reader = new StreamReader(input, Encoding.UTF8);
                string xml = rewrite(reader.ReadToEnd());
                using var writer = new StreamWriter(output, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write(xml);
            }

            Assert.True(found, $"Part ending in '{suffix}' was not found.");
        }

        return result.ToArray();
    }

    private static byte[] RelocatePersonsAndOccupyCanonicalPath(byte[] package)
    {
        const string Canonical = "xl/persons/person.xml";
        const string Relocated = "xl/persons/authors.xml";

        using var sourceStream = new MemoryStream(package);
        using var source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
        using var result = new MemoryStream();
        using (var target = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                string name = entry.FullName.Equals(Canonical, StringComparison.OrdinalIgnoreCase)
                    ? Relocated
                    : entry.FullName;
                ZipArchiveEntry copy = target.CreateEntry(name, CompressionLevel.Fastest);
                using Stream output = copy.Open();
                using Stream input = entry.Open();

                if (entry.FullName.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase))
                {
                    XDocument relationships = XDocument.Load(input);
                    XNamespace ns = relationships.Root!.Name.Namespace;
                    XElement persons = Assert.Single(
                        relationships.Root.Elements(),
                        element => element.Attribute("Type")?.Value == XlsxSchema.RelPersons);
                    persons.SetAttributeValue("Target", "persons/authors.xml");
                    relationships.Root.Add(new XElement(
                        ns + "Relationship",
                        new XAttribute("Id", "rIdCanonicalPersonsProbe"),
                        new XAttribute("Type", "urn:cellwright:test/canonical-persons-probe"),
                        new XAttribute("Target", "persons/person.xml")));
                    relationships.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
                    continue;
                }

                if (entry.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                {
                    XDocument contentTypes = XDocument.Load(input);
                    XElement persons = Assert.Single(
                        contentTypes.Root!.Elements(),
                        element => element.Attribute("PartName")?.Value == "/xl/persons/person.xml");
                    persons.SetAttributeValue("PartName", "/xl/persons/authors.xml");
                    contentTypes.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
                    continue;
                }

                input.CopyTo(output);
            }

            ZipArchiveEntry occupied = target.CreateEntry(Canonical, CompressionLevel.Fastest);
            using var writer = new StreamWriter(
                occupied.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            writer.Write(CanonicalPersonsProbe);
        }

        return result.ToArray();
    }

    private static string DuplicateFirstCommentAt(string xml, string from, string to)
    {
        int start = xml.IndexOf("<comment ref=", StringComparison.Ordinal);
        int end = xml.IndexOf("</comment>", start, StringComparison.Ordinal);
        Assert.True(start >= 0 && end >= 0);
        end += "</comment>".Length;
        string copy = xml[start..end].Replace(
            $"ref=\"{from}\"", $"ref=\"{to}\"", StringComparison.Ordinal);
        int insert = xml.IndexOf("</commentList>", end, StringComparison.Ordinal);
        Assert.True(insert >= 0);
        return xml.Insert(insert, copy);
    }

    private static async Task<Workbook> LoadAsync(byte[] package)
    {
        using var stream = new MemoryStream(package);
        return await Workbook.LoadAsync(
            stream, new LoadOptions(), TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task AThreadSurvivesASaveAndReload()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet["A1"] = "Revenue";

        Person author = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Is this net or gross?", author);

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);
        Worksheet read = reloaded.Sheets[0];

        ThreadedComment comment = Assert.Single(read.ThreadedComments);
        Assert.Equal("Is this net or gross?", comment.Text);
        Assert.Equal(CellRef.Parse("A1"), comment.Cell);
        Assert.True(comment.IsFirst);

        Person person = Assert.Single(reloaded.People);
        Assert.Equal("Nadia", person.DisplayName);
        Assert.Equal(person.Id, comment.PersonId);
    }

    [Fact]
    public async Task ARepliesToItAndTheThreadReadsAsAConversation()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        Person asker = book.People.Add("Nadia");
        Person answerer = book.People.Add("Tom");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("B2"), "Net or gross?", asker);
        sheet.ThreadedComments.Reply(first, "Net.", answerer);

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);
        Worksheet read = reloaded.Sheets[0];

        Assert.Equal(2, read.ThreadedComments.Count);
        Assert.Equal(2, reloaded.People.Count);

        ThreadedComment thread = Assert.Single(read.ThreadedComments.Threads);
        Assert.Equal("Net or gross?", thread.Text);

        ThreadedComment reply = Assert.Single(thread.Replies);
        Assert.Equal("Net.", reply.Text);
        Assert.Equal(thread.Id, reply.ParentId);
        Assert.Equal(CellRef.Parse("B2"), reply.Cell);
    }

    [Fact]
    public async Task AReplyToAReplyStaysInTheSameThread()
    {
        // The format has one level of nesting and Excel shows a flat conversation, so answering
        // a reply must answer the thread rather than build a chain.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", person);
        ThreadedComment second = sheet.ThreadedComments.Reply(first, "Two", person);
        sheet.ThreadedComments.Reply(second, "Three", person);

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        ThreadedComment thread = Assert.Single(reloaded.Sheets[0].ThreadedComments.Threads);
        Assert.Equal(2, thread.Replies.Count);
    }

    [Fact]
    public void AuthorsMustComeFromTheOwningWorkbook()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person own = book.People.Add("Nadia");
        Person unattached = Person.Named("Nadia");
        Person other = Workbook.Create().People.Add("Tom");

        Assert.Throws<ArgumentException>(() =>
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Unattached", unattached));
        Assert.Throws<ArgumentException>(() =>
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Other workbook", other));

        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Own", own);
        Assert.Throws<ArgumentException>(() => sheet.ThreadedComments.Reply(first, "Other", other));
        Assert.Single(sheet.ThreadedComments);
    }

    [Fact]
    public void AReplyParentMustComeFromTheSameCollection()
    {
        Workbook book = Workbook.Create();
        Worksheet firstSheet = book.Sheets.Add("First");
        Worksheet secondSheet = book.Sheets.Add("Second");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = firstSheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", person);

        Assert.Throws<ArgumentException>(() =>
            secondSheet.ThreadedComments.Reply(first, "Cross-sheet", person));

        Assert.True(firstSheet.ThreadedComments.Remove(first));
        Assert.Throws<ArgumentException>(() =>
            firstSheet.ThreadedComments.Reply(first, "Removed", person));
        Assert.Empty(secondSheet.ThreadedComments);
    }

    [Fact]
    public void DetachedCollectionsKeepTheirLegacyAuthoringBehavior()
    {
        var comments = new ThreadedCommentCollection();
        Person person = Person.Named("Nadia");

        ThreadedComment first = comments.Add(CellRef.Parse("A1"), "One", person);
        comments.Reply(first, "Two", person);

        Assert.Equal(2, comments.Count);
        Assert.Single(Assert.Single(comments.Threads).Replies);
    }

    [Fact]
    public void ReadingThreadsTwiceDoesNotDuplicateReplies()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", person);
        sheet.ThreadedComments.Reply(first, "Two", person);

        ThreadedComment firstProjection = Assert.Single(sheet.ThreadedComments.Threads);
        Assert.Single(firstProjection.Replies);

        ThreadedComment secondProjection = Assert.Single(sheet.ThreadedComments.Threads);
        Assert.Same(firstProjection, secondProjection);
        Assert.Single(secondProjection.Replies);
    }

    [Fact]
    public async Task AThreadGetsALegacyNoteBesideItForOlderReaders()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        ThreadedComment thread =
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Check this", book.People.Add("Nadia"));

        (_, byte[] bytes) = await SaveAndReloadAsync(book);

        // The note names the thread it stands for, which is how Excel knows to hide it.
        string comments = PartOf(bytes, "comments1.xml");
        Assert.Contains("tc=", comments, StringComparison.Ordinal);
        Assert.Contains($"xr:uid=\"{thread.Id}\"", comments, StringComparison.Ordinal);
        Assert.Contains("Check this", comments, StringComparison.Ordinal);

        // And the VML gives it a box, without which nothing on the cell is visible at all.
        Assert.True(HasPart(bytes, ".vml"));
    }

    [Fact]
    public async Task ThePlaceholderNoteDoesNotComeBackAsANote()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Check this", book.People.Add("Nadia"));

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        // It is a stand-in for the thread, not a note anybody wrote.
        Assert.Empty(reloaded.Sheets[0].Comments);
        Assert.Single(reloaded.Sheets[0].ThreadedComments);
    }

    [Fact]
    public async Task AnOrdinaryAuthorWhoseNameStartsWithTcIsNotDiscarded()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Comments.Add(CellRef.Parse("A1"), "An ordinary note", "tc=Reviewer");

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        CellComment note = Assert.Single(reloaded.Sheets[0].Comments);
        Assert.Equal("tc=Reviewer", note.Author);
        Assert.Equal("An ordinary note", note.Text);
        Assert.DoesNotContain(
            reloaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "UnmatchedThreadedCommentPlaceholder");
    }

    [Fact]
    public async Task PlaceholderAuthorTokenIsTheCorrespondenceKeyAndUidMismatchIsRepaired()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        ThreadedComment thread =
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Thread", book.People.Add("Nadia"));
        (_, byte[] original) = await SaveAndReloadAsync(book);
        string otherId = Person.NewId();
        byte[] malformed = RewritePart(original, "comments1.xml", xml => xml.Replace(
            $"xr:uid=\"{thread.Id}\"", $"xr:uid=\"{otherId}\"", StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(malformed);

        Assert.Single(loaded.Sheets[0].ThreadedComments);
        Assert.Empty(loaded.Sheets[0].Comments);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentPlaceholderUidMismatch");

        (_, byte[] repaired) = await SaveAndReloadAsync(loaded);
        Assert.Contains(
            $"xr:uid=\"{thread.Id}\"",
            SheetRelatedPart(repaired, 1, "/comments"),
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task PlaceholderCellMovesEveryMessageInTheThread()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment root = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", person);
        sheet.ThreadedComments.Reply(root, "Two", person);
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] moved = RewritePart(original, "comments1.xml", static xml => xml.Replace(
            "<comment ref=\"A1\"", "<comment ref=\"B2\"", StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(moved);

        Assert.All(loaded.Sheets[0].ThreadedComments, static comment =>
            Assert.Equal(CellRef.Parse("B2"), comment.Cell));
        Assert.Empty(loaded.Sheets[0].Comments);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentMovedToPlaceholder");
    }

    [Fact]
    public async Task MissingAndOrphanPlaceholderSidesAreRemovedWithDiagnostics()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        ThreadedComment thread =
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Thread", book.People.Add("Nadia"));
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] disconnected = RewritePart(original, "comments1.xml", xml => xml.Replace(
            $">tc={thread.Id}<", $">tc={Person.NewId()}<", StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(disconnected);

        Assert.Empty(loaded.Sheets[0].ThreadedComments);
        Assert.Empty(loaded.Sheets[0].Comments);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentWithoutPlaceholderRemoved");
        Assert.Contains(
            loaded.LoadDiagnostics.Warnings,
            static warning => warning.Code == "OrphanThreadedCommentPlaceholderRemoved");
    }

    [Fact]
    public async Task APlaceholderOnAnotherSheetMovesTheWholeThreadAndKeepsOpaqueMetadata()
    {
        Workbook book = Workbook.Create();
        Worksheet firstSheet = book.Sheets.Add("First");
        Worksheet secondSheet = book.Sheets.Add("Second");
        Person person = book.People.Add("Nadia");
        ThreadedComment root = firstSheet.ThreadedComments.Add(CellRef.Parse("A1"), "Ask @Nadia", person);
        firstSheet.ThreadedComments.Reply(root, "Reply", person);
        secondSheet.Comments.Add(CellRef.Parse("C3"), "Compatibility note", "Reviewer");
        (_, byte[] original) = await SaveAndReloadAsync(book);
        const string collectionExtension = "{55555555-5555-5555-5555-555555555555}";

        byte[] withMetadata = RewritePart(original, "threadedComment1.xml", xml => xml
            .Replace(
                "</text>",
                $"</text><mentions><mention mentionpersonId=\"{person.Id}\" mentionId=\"{Person.NewId()}\" " +
                "startIndex=\"4\" length=\"6\"/></mentions>",
                StringComparison.Ordinal)
            .Replace(
                "</ThreadedComments>",
                $"<x:extLst><x:ext uri=\"{collectionExtension}\"><probe xmlns=\"urn:cellwright:test\"/>" +
                "</x:ext></x:extLst></ThreadedComments>",
                StringComparison.Ordinal));
        byte[] withoutLocalPlaceholder = RewritePart(withMetadata, "comments1.xml", xml => xml.Replace(
            $">tc={root.Id}<", ">Reviewer<", StringComparison.Ordinal));
        byte[] crossSheet = RewritePart(withoutLocalPlaceholder, "comments2.xml", xml => xml.Replace(
            ">Reviewer<", $">tc={root.Id}<", StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(crossSheet);

        Assert.Empty(loaded.Sheets["First"].ThreadedComments);
        Assert.Equal(2, loaded.Sheets["Second"].ThreadedComments.Count);
        Assert.Null(loaded.Sheets["First"].ThreadedComments.ExtensionListXml);
        Assert.Contains(
            collectionExtension,
            Assert.IsType<string>(loaded.Sheets["Second"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        ThreadedComment moved = Assert.Single(loaded.Sheets["Second"].ThreadedComments.Threads);
        Assert.Equal(CellRef.Parse("C3"), moved.Cell);
        Assert.True(moved.HasOpaqueMetadata);
        ThreadedComment reply = Assert.Single(moved.Replies);
        Assert.Equal(moved.Id, reply.ParentId);
        Assert.All(loaded.Sheets["Second"].ThreadedComments, static comment =>
            Assert.Equal(CellRef.Parse("C3"), comment.Cell));
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentMovedToPlaceholder");

        (Workbook reloaded, _) = await SaveAndReloadAsync(loaded);
        ThreadedComment roundTrip = Assert.Single(reloaded.Sheets["Second"].ThreadedComments.Threads);
        Assert.True(roundTrip.HasOpaqueMetadata);
        Assert.Contains(
            collectionExtension,
            Assert.IsType<string>(reloaded.Sheets["Second"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        Assert.Equal(roundTrip.Id, Assert.Single(roundTrip.Replies).ParentId);
        Assert.Empty(reloaded.Sheets["First"].ThreadedComments);
    }

    [Fact]
    public async Task AdditionalPlaceholderCopiesConversationWithFreshRemappedIds()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment root = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Ask @Nadia", person);
        sheet.ThreadedComments.Reply(root, "Reply", person);
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] withMetadata = RewritePart(original, "threadedComment1.xml", xml => xml.Replace(
            "</text>",
            $"</text><mentions><mention mentionpersonId=\"{person.Id}\" mentionId=\"{Person.NewId()}\" " +
            "startIndex=\"4\" length=\"6\"/></mentions>",
            StringComparison.Ordinal));
        byte[] duplicated = RewritePart(
            withMetadata,
            "comments1.xml",
            static xml => DuplicateFirstCommentAt(xml, "A1", "B2"));

        Workbook loaded = await LoadAsync(duplicated);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentCopiedForPlaceholder");
        AssertConversationsAreIndependentCopies(loaded.Sheets[0]);

        (Workbook reloaded, _) = await SaveAndReloadAsync(loaded);
        AssertConversationsAreIndependentCopies(reloaded.Sheets[0]);

        static void AssertConversationsAreIndependentCopies(Worksheet worksheet)
        {
            Assert.Equal(4, worksheet.ThreadedComments.Count);
            IReadOnlyList<ThreadedComment> roots = worksheet.ThreadedComments.Threads;
            Assert.Equal(2, roots.Count);
            Assert.Equal(new[] { CellRef.Parse("A1"), CellRef.Parse("B2") }, roots.Select(static item => item.Cell));
            Assert.All(roots, static item => Assert.True(item.HasOpaqueMetadata));
            Assert.NotEqual(roots[0].Id, roots[1].Id);
            Assert.Equal(
                worksheet.ThreadedComments.Count,
                worksheet.ThreadedComments.Select(static comment => comment.Id)
                    .Distinct(StringComparer.OrdinalIgnoreCase).Count());
            foreach (ThreadedComment item in roots)
            {
                ThreadedComment reply = Assert.Single(item.Replies);
                Assert.Equal(item.Id, reply.ParentId);
                Assert.Equal(item.Cell, reply.Cell);
            }
        }
    }

    [Fact]
    public async Task CrossSheetCollectionMetadataConflictKeepsBothListsAndWarns()
    {
        Workbook book = Workbook.Create();
        Worksheet firstSheet = book.Sheets.Add("First");
        Worksheet secondSheet = book.Sheets.Add("Second");
        Person person = book.People.Add("Nadia");
        ThreadedComment moving = firstSheet.ThreadedComments.Add(CellRef.Parse("A1"), "Move", person);
        secondSheet.ThreadedComments.Add(CellRef.Parse("B2"), "Stay", person);
        secondSheet.Comments.Add(CellRef.Parse("C3"), "Destination", "Transfer");
        (_, byte[] original) = await SaveAndReloadAsync(book);
        const string firstExtension = "{66666666-6666-6666-6666-666666666666}";
        const string secondExtension = "{77777777-7777-7777-7777-777777777777}";

        byte[] firstMarkup = RewritePart(original, "threadedComment1.xml", xml => xml.Replace(
            "</ThreadedComments>",
            $"<x:extLst><x:ext uri=\"{firstExtension}\"/></x:extLst></ThreadedComments>",
            StringComparison.Ordinal));
        byte[] bothMarkup = RewritePart(firstMarkup, "threadedComment2.xml", xml => xml.Replace(
            "</ThreadedComments>",
            $"<x:extLst><x:ext uri=\"{secondExtension}\"/></x:extLst></ThreadedComments>",
            StringComparison.Ordinal));
        byte[] withoutSourcePlaceholder = RewritePart(bothMarkup, "comments1.xml", xml => xml.Replace(
            $">tc={moving.Id}<", ">Reviewer<", StringComparison.Ordinal));
        byte[] conflict = RewritePart(withoutSourcePlaceholder, "comments2.xml", xml => xml.Replace(
            ">Transfer<", $">tc={moving.Id}<", StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(conflict);

        Assert.Empty(loaded.Sheets["First"].ThreadedComments);
        Assert.Equal(2, loaded.Sheets["Second"].ThreadedComments.Threads.Count);
        Assert.Contains(
            firstExtension,
            Assert.IsType<string>(loaded.Sheets["First"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        Assert.Contains(
            secondExtension,
            Assert.IsType<string>(loaded.Sheets["Second"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentCollectionMetadataConflict");

        (Workbook reloaded, byte[] rewritten) = await SaveAndReloadAsync(loaded);
        Assert.Empty(reloaded.Sheets["First"].ThreadedComments);
        Assert.Equal(2, reloaded.Sheets["Second"].ThreadedComments.Threads.Count);
        Assert.Contains(
            firstExtension,
            Assert.IsType<string>(reloaded.Sheets["First"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        Assert.Contains(
            secondExtension,
            Assert.IsType<string>(reloaded.Sheets["Second"].ThreadedComments.ExtensionListXml),
            StringComparison.Ordinal);
        IReadOnlyList<string> threadedParts = XmlPartsUnder(rewritten, "threadedComments/");
        Assert.Equal(2, threadedParts.Count);
        Assert.Single(threadedParts, part =>
            part.Contains(firstExtension, StringComparison.Ordinal) &&
            !part.Contains("<threadedComment ", StringComparison.Ordinal));
        Assert.Single(threadedParts, part => part.Contains(secondExtension, StringComparison.Ordinal));
    }

    [Fact]
    public async Task NotesAndThreadsLiveTogetherOnOneSheet()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.Comments.Add(CellRef.Parse("A1"), "An ordinary note", "Reviewer");
        sheet.ThreadedComments.Add(CellRef.Parse("B2"), "A thread", book.People.Add("Nadia"));

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);
        Worksheet read = reloaded.Sheets[0];

        CellComment note = Assert.Single(read.Comments);
        Assert.Equal("An ordinary note", note.Text);
        Assert.Equal(CellRef.Parse("A1"), note.Cell);

        ThreadedComment thread = Assert.Single(read.ThreadedComments);
        Assert.Equal(CellRef.Parse("B2"), thread.Cell);
    }

    [Fact]
    public async Task EditingANoteNoLongerDetachesTheThreads()
    {
        // The old behaviour: rewriting the comments part dropped the placeholders and cut every
        // thread loose. Now the placeholders are regenerated from the threads.
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "A thread", book.People.Add("Nadia"));

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        // Touch the notes of the reloaded book, which is what used to lose the threads.
        reloaded.Sheets[0].Comments.Add(CellRef.Parse("C3"), "A new note", "Someone");

        (Workbook again, _) = await SaveAndReloadAsync(reloaded);

        Assert.Single(again.Sheets[0].ThreadedComments);
        Assert.Equal("A thread", again.Sheets[0].ThreadedComments[0].Text);
        Assert.Single(again.Sheets[0].Comments);
    }

    [Fact]
    public async Task ResolvingAThreadSurvives()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        ThreadedComment comment =
            sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Sorted?", book.People.Add("Nadia"));
        comment.Resolved = true;

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        Assert.True(reloaded.Sheets[0].ThreadedComments[0].Resolved);
    }

    [Fact]
    public async Task EditingTextOnALoadedThreadSurvivesAnotherSave()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Before", book.People.Add("Nadia"));
        (Workbook loaded, _) = await SaveAndReloadAsync(book);

        loaded.Sheets[0].ThreadedComments[0].Text = "After";
        (Workbook reloaded, _) = await SaveAndReloadAsync(loaded);

        Assert.Equal("After", reloaded.Sheets[0].ThreadedComments[0].Text);
    }

    [Fact]
    public async Task ResolvingALoadedThreadSurvivesAnotherSave()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Sorted?", book.People.Add("Nadia"));
        (Workbook loaded, _) = await SaveAndReloadAsync(book);

        loaded.Sheets[0].ThreadedComments[0].Resolved = true;
        (Workbook reloaded, _) = await SaveAndReloadAsync(loaded);

        Assert.True(reloaded.Sheets[0].ThreadedComments[0].Resolved);
    }

    [Fact]
    public async Task MovingALoadedThreadSurvivesAnotherSave()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Move me", book.People.Add("Nadia"));
        (Workbook loaded, _) = await SaveAndReloadAsync(book);

        loaded.Sheets[0].ThreadedComments[0].Cell = CellRef.Parse("C3");
        (Workbook reloaded, _) = await SaveAndReloadAsync(loaded);

        Assert.Equal(CellRef.Parse("C3"), reloaded.Sheets[0].ThreadedComments[0].Cell);
    }

    [Fact]
    public async Task StructuralEditsMoveThreadedCommentsWithTheirCells()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("B2"), "One", person);
        sheet.ThreadedComments.Reply(first, "Two", person);

        sheet.InsertRows(2, 2);    // B2 -> B4
        sheet.InsertColumns(2);    // B4 -> C4
        sheet.DeleteRows(3);       // C4 -> C3
        sheet.DeleteColumns(2);    // C3 -> B3

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        Assert.All(reloaded.Sheets[0].ThreadedComments,
            comment => Assert.Equal(CellRef.Parse("B3"), comment.Cell));
    }

    [Fact]
    public async Task StructuralEditOutsideTheThreadKeepsItsPartInCopyOnWriteMode()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("B2"), "One", book.People.Add("Nadia"));
        (Workbook loaded, _) = await SaveAndReloadAsync(book);

        Assert.False(loaded.Sheets[0].ThreadedComments.IsModified);

        loaded.Sheets[0].InsertRows(20);

        Assert.False(loaded.Sheets[0].ThreadedComments.IsModified);
    }

    [Fact]
    public async Task DeletingTheAnchorCellDropsTheWholeThread()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("B2"), "One", person);
        sheet.ThreadedComments.Reply(first, "Two", person);

        sheet.DeleteRows(2);
        (Workbook reloaded, byte[] bytes) = await SaveAndReloadAsync(book);

        Assert.Empty(reloaded.Sheets[0].ThreadedComments);
        Assert.False(HasPart(bytes, "threadedComment1.xml"));
    }

    [Fact]
    public async Task RemovingAThreadTakesItsRepliesWithIt()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", person);
        sheet.ThreadedComments.Reply(first, "Two", person);

        sheet.ThreadedComments.Remove(first);

        (Workbook reloaded, byte[] bytes) = await SaveAndReloadAsync(book);

        Assert.Empty(reloaded.Sheets[0].ThreadedComments);
        Assert.False(HasPart(bytes, "threadedComment1.xml"));
    }

    [Fact]
    public async Task TheSameAuthorIsOnePerson()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");

        Person once = book.People.Add("Nadia");
        Person twice = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", once);
        sheet.ThreadedComments.Add(CellRef.Parse("A2"), "Two", twice);

        (Workbook reloaded, _) = await SaveAndReloadAsync(book);

        Assert.Single(reloaded.People);
        Assert.Equal(once.Id, twice.Id);
    }

    [Fact]
    public async Task PeopleRegenerationReusesItsSourcePartWithoutOverwritingAnOccupiedCanonicalPath()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "One", book.People.Add("Nadia"));
        (_, byte[] original) = await SaveAndReloadAsync(book);

        Workbook loaded = await LoadAsync(RelocatePersonsAndOccupyCanonicalPath(original));
        loaded.People.Add("Tom");

        (Workbook reloaded, byte[] rewritten) = await SaveAndReloadAsync(loaded);

        Assert.Contains(reloaded.People, static person => person.DisplayName == "Nadia");
        Assert.Contains(reloaded.People, static person => person.DisplayName == "Tom");
        using var archive = new ZipArchive(new MemoryStream(rewritten), ZipArchiveMode.Read);
        Assert.Single(
            archive.Entries,
            static entry => entry.FullName.Equals("xl/persons/person.xml", StringComparison.OrdinalIgnoreCase));
        using (var reader = new StreamReader(archive.GetEntry("xl/persons/person.xml")!.Open(), Encoding.UTF8))
            Assert.Equal(CanonicalPersonsProbe, reader.ReadToEnd());
        using (var reader = new StreamReader(archive.GetEntry("xl/persons/authors.xml")!.Open(), Encoding.UTF8))
        {
            string people = reader.ReadToEnd();
            Assert.Contains("displayName=\"Nadia\"", people, StringComparison.Ordinal);
            Assert.Contains("displayName=\"Tom\"", people, StringComparison.Ordinal);
        }

        XDocument relationships;
        using (Stream stream = archive.GetEntry("xl/_rels/workbook.xml.rels")!.Open())
            relationships = XDocument.Load(stream);
        XElement[] links = [.. relationships.Root!.Elements()];
        Assert.Equal(
            "persons/authors.xml",
            Assert.Single(links, element => element.Attribute("Type")?.Value == XlsxSchema.RelPersons)
                .Attribute("Target")?.Value);
        Assert.Equal(
            "persons/person.xml",
            Assert.Single(
                    links,
                    element => element.Attribute("Type")?.Value ==
                               "urn:cellwright:test/canonical-persons-probe")
                .Attribute("Target")?.Value);
    }

    [Fact]
    public async Task AWorkbookWithoutThreadsWritesNeitherPart()
    {
        Workbook book = Workbook.Create();
        book.Sheets.Add("Sheet1")["A1"] = 1;

        (_, byte[] bytes) = await SaveAndReloadAsync(book);

        Assert.False(HasPart(bytes, "threadedComment1.xml"));
        Assert.False(HasPart(bytes, "person.xml"));
    }

    [Fact]
    public async Task TheThreadNamesTheCellAndTheAuthorTheFormatExpects()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("C7"), "Text", person);

        (_, byte[] bytes) = await SaveAndReloadAsync(book);

        string part = PartOf(bytes, "threadedComment1.xml");
        Assert.Contains("ref=\"C7\"", part, StringComparison.Ordinal);
        Assert.Contains($"personId=\"{person.Id}\"", part, StringComparison.Ordinal);
        Assert.Contains("<text>Text</text>", part, StringComparison.Ordinal);

        string people = PartOf(bytes, "person.xml");
        Assert.Contains("displayName=\"Nadia\"", people, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MentionsAndExtensionListsSurviveRewriteAndSheetCopy()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        Person nadia = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("B2"), "Ask @Nadia", nadia);

        (_, byte[] original) = await SaveAndReloadAsync(book);
        string mentionId = Person.NewId();
        const string messageExtension = "{11111111-1111-1111-1111-111111111111}";
        const string listExtension = "{22222222-2222-2222-2222-222222222222}";
        const string personExtension = "{33333333-3333-3333-3333-333333333333}";
        const string peopleExtension = "{44444444-4444-4444-4444-444444444444}";

        byte[] withCommentMarkup = RewritePart(original, "threadedComment1.xml", xml => xml
            .Replace("</text>",
                $"</text><mentions><mention mentionpersonId=\"{nadia.Id}\" mentionId=\"{mentionId}\" startIndex=\"4\" length=\"6\"/></mentions>" +
                $"<x:extLst><x:ext uri=\"{messageExtension}\"><probe xmlns=\"urn:cellwright:test\" kind=\"message\"/></x:ext></x:extLst>",
                StringComparison.Ordinal)
            .Replace("</ThreadedComments>",
                $"<x:extLst><x:ext uri=\"{listExtension}\"><probe xmlns=\"urn:cellwright:test\" kind=\"list\"/></x:ext></x:extLst></ThreadedComments>",
                StringComparison.Ordinal));

        byte[] withAllMarkup = RewritePart(withCommentMarkup, "person.xml", xml => xml
            .Replace(" providerId=\"None\"/>",
                $" providerId=\"None\"><x:extLst><x:ext uri=\"{personExtension}\"><probe xmlns=\"urn:cellwright:test\" kind=\"person\"/></x:ext></x:extLst></person>",
                StringComparison.Ordinal)
            .Replace("</personList>",
                $"<x:extLst><x:ext uri=\"{peopleExtension}\"><probe xmlns=\"urn:cellwright:test\" kind=\"people\"/></x:ext></x:extLst></personList>",
                StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(withAllMarkup);
        Worksheet source = loaded.Sheets["Report"];
        source.ThreadedComments[0].Resolved = true;
        loaded.Sheets.Copy(source, "Report copy");
        loaded.People.Add("Tom");

        (Workbook reloaded, byte[] rewritten) = await SaveAndReloadAsync(loaded);

        Assert.Equal(2, reloaded.Sheets.Count);
        foreach (int sheetIndex in new[] { 1, 2 })
        {
            string comments = SheetRelatedPart(rewritten, sheetIndex, "/threadedComment");
            Assert.Contains($"mentionpersonId=\"{nadia.Id}\"", comments, StringComparison.Ordinal);
            Assert.Contains($"mentionId=\"{mentionId}\"", comments, StringComparison.Ordinal);
            Assert.Contains($"uri=\"{messageExtension}\"", comments, StringComparison.Ordinal);
            Assert.Contains($"uri=\"{listExtension}\"", comments, StringComparison.Ordinal);
        }

        string people = PartOf(rewritten, "person.xml");
        Assert.Contains($"uri=\"{personExtension}\"", people, StringComparison.Ordinal);
        Assert.Contains($"uri=\"{peopleExtension}\"", people, StringComparison.Ordinal);
    }

    [Fact]
    public async Task EditingTextDropsStaleMessageMetadataAndWarns()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        Person nadia = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("B2"), "Ask @Nadia", nadia);

        (_, byte[] original) = await SaveAndReloadAsync(book);
        const string extension = "{11111111-1111-1111-1111-111111111111}";
        byte[] withMarkup = RewritePart(original, "threadedComment1.xml", xml => xml.Replace(
            "</text>",
            $"</text><mentions><mention mentionpersonId=\"{nadia.Id}\" mentionId=\"{Person.NewId()}\" " +
            "startIndex=\"4\" length=\"6\"/></mentions>" +
            $"<x:extLst><x:ext uri=\"{extension}\"><probe xmlns=\"urn:cellwright:test\"/></x:ext></x:extLst>",
            StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(withMarkup);
        var warnings = new List<WorkbookWarning>();
        loaded.OnWarning = warnings.Add;

        loaded.Sheets[0].ThreadedComments[0].Text = "Ask somebody";

        WorkbookWarning warning = Assert.Single(warnings);
        Assert.Equal("ThreadedCommentTextMetadataDropped", warning.Code);
        Assert.Equal(CellRef.Parse("B2"), warning.Cell);

        (_, byte[] rewritten) = await SaveAndReloadAsync(loaded);
        string comments = SheetRelatedPart(rewritten, 1, "/threadedComment");
        Assert.DoesNotContain("<mentions", comments, StringComparison.Ordinal);
        Assert.DoesNotContain(extension, comments, StringComparison.Ordinal);
        Assert.Contains("<text>Ask somebody</text>", comments, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OpaqueMessageMetadataCanBeClearedExplicitlyWithoutAnEditWarning()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Report");
        Person nadia = book.People.Add("Nadia");
        sheet.ThreadedComments.Add(CellRef.Parse("B2"), "Ask @Nadia", nadia);
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] withMention = RewritePart(original, "threadedComment1.xml", xml => xml.Replace(
            "</text>",
            $"</text><mentions><mention mentionpersonId=\"{nadia.Id}\" mentionId=\"{Person.NewId()}\" " +
            "startIndex=\"4\" length=\"6\"/></mentions>",
            StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(withMention);
        ThreadedComment comment = loaded.Sheets[0].ThreadedComments[0];
        var warnings = new List<WorkbookWarning>();
        loaded.OnWarning = warnings.Add;

        Assert.True(comment.HasOpaqueMetadata);
        Assert.True(comment.ClearOpaqueMetadata());
        Assert.False(comment.HasOpaqueMetadata);
        Assert.False(comment.ClearOpaqueMetadata());
        Assert.Empty(warnings);

        (_, byte[] rewritten) = await SaveAndReloadAsync(loaded);
        string onlyThreadPart = Assert.Single(XmlPartsUnder(rewritten, "threadedComments/"));
        Assert.DoesNotContain(
            "<mentions",
            onlyThreadPart,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingCreatedTimeStaysUnknownWhenThePartIsRewritten()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Unknown provenance", book.People.Add("Nadia"));
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] withoutDate = RewritePart(
            original,
            "threadedComment1.xml",
            static xml => Regex.Replace(xml, " dT=\"[^\"]*\"", string.Empty));

        Workbook loaded = await LoadAsync(withoutDate);
        Assert.Null(loaded.Sheets[0].ThreadedComments[0].Created);
        loaded.Sheets[0].ThreadedComments[0].Resolved = true;

        (Workbook reloaded, byte[] rewritten) = await SaveAndReloadAsync(loaded);
        Assert.DoesNotContain(
            " dT=",
            SheetRelatedPart(rewritten, 1, "/threadedComment"),
            StringComparison.Ordinal);
        Assert.Null(reloaded.Sheets[0].ThreadedComments[0].Created);
    }

    [Fact]
    public async Task DuplicateIdsAreDiagnosedAndReconciliationKeepsTheFirstOwner()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "First", person);
        ThreadedComment duplicate = sheet.ThreadedComments.Add(CellRef.Parse("A2"), "Duplicate", person);
        sheet.ThreadedComments.Reply(first, "Reply", person);
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] malformed = RewritePart(original, "threadedComment1.xml", xml => xml.Replace(
            $" id=\"{duplicate.Id}\"",
            $" id=\"{first.Id}\"",
            StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(malformed);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "DuplicateThreadedCommentId");

        ThreadedComment root = Assert.Single(loaded.Sheets[0].ThreadedComments.Threads);
        Assert.Equal("First", root.Text);
        Assert.Equal("Reply", Assert.Single(root.Replies).Text);
        Assert.DoesNotContain(loaded.Sheets[0].ThreadedComments, static comment => comment.Text == "Duplicate");

        ThreadedComment secondProjection = Assert.Single(loaded.Sheets[0].ThreadedComments.Threads);
        Assert.Equal("Reply", Assert.Single(secondProjection.Replies).Text);
    }

    [Fact]
    public void SheetCopyKeepsTheFirstOwnerOfAMalformedDuplicateId()
    {
        Workbook book = Workbook.Create();
        Worksheet source = book.Sheets.Add("Source");
        Person person = book.People.Add("Nadia");
        string rootId = Person.NewId();
        string duplicateId = Person.NewId();

        source.ThreadedComments.AddFromFile(new ThreadedComment
        {
            Cell = CellRef.Parse("A1"),
            Text = "Root",
            PersonId = person.Id,
            Id = rootId,
        });
        source.ThreadedComments.AddFromFile(new ThreadedComment
        {
            Cell = CellRef.Parse("A1"),
            Text = "First duplicate",
            PersonId = person.Id,
            Id = duplicateId,
            ParentId = rootId,
        });
        source.ThreadedComments.AddFromFile(new ThreadedComment
        {
            Cell = CellRef.Parse("A1"),
            Text = "Second duplicate",
            PersonId = person.Id,
            Id = duplicateId,
            ParentId = rootId,
        });
        source.ThreadedComments.AddFromFile(new ThreadedComment
        {
            Cell = CellRef.Parse("A1"),
            Text = "Nested reply",
            PersonId = person.Id,
            Id = Person.NewId(),
            ParentId = duplicateId,
        });

        Worksheet copy = book.Sheets.Copy(source, "Copy");

        ThreadedComment copiedRoot = Assert.Single(copy.ThreadedComments.Threads);
        Assert.Equal(
            new[] { "First duplicate", "Second duplicate" },
            copiedRoot.Replies.Select(static reply => reply.Text));
        ThreadedComment firstDuplicate = Assert.Single(
            copy.ThreadedComments,
            static comment => comment.Text == "First duplicate");
        ThreadedComment nestedReply = Assert.Single(
            copy.ThreadedComments,
            static comment => comment.Text == "Nested reply");
        Assert.Equal(firstDuplicate.Id, nestedReply.ParentId);
        Assert.Equal("Nested reply", Assert.Single(firstDuplicate.Replies).Text);
    }

    [Fact]
    public async Task ParentCyclesAreDiagnosedAndEveryCycleMemberBecomesARoot()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        Person person = book.People.Add("Nadia");
        ThreadedComment first = sheet.ThreadedComments.Add(CellRef.Parse("A1"), "First", person);
        ThreadedComment second = sheet.ThreadedComments.Add(CellRef.Parse("A2"), "Second", person);
        (_, byte[] original) = await SaveAndReloadAsync(book);
        byte[] malformed = RewritePart(original, "threadedComment1.xml", xml => xml
            .Replace(
                $" id=\"{first.Id}\"",
                $" id=\"{first.Id}\" parentId=\"{second.Id}\"",
                StringComparison.Ordinal)
            .Replace(
                $" id=\"{second.Id}\"",
                $" id=\"{second.Id}\" parentId=\"{first.Id}\"",
                StringComparison.Ordinal));

        Workbook loaded = await LoadAsync(malformed);
        Assert.Contains(
            loaded.LoadDiagnostics!.Warnings,
            static warning => warning.Code == "ThreadedCommentParentCycle");

        IReadOnlyList<ThreadedComment> roots = loaded.Sheets[0].ThreadedComments.Threads;
        Assert.Equal(new[] { "First", "Second" }, roots.Select(static root => root.Text));
        Assert.All(roots, static root => Assert.Empty(root.Replies));
    }
}
