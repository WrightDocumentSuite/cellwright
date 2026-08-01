using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Threaded comments: the conversation attached to a cell, the people who wrote it, and the
/// legacy note that stands beside it for readers that predate threads.
/// </summary>
public sealed class ThreadedCommentTests
{
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
    public async Task AThreadGetsALegacyNoteBesideItForOlderReaders()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Sheet1");
        sheet.ThreadedComments.Add(CellRef.Parse("A1"), "Check this", book.People.Add("Nadia"));

        (_, byte[] bytes) = await SaveAndReloadAsync(book);

        // The note names the thread it stands for, which is how Excel knows to hide it.
        string comments = PartOf(bytes, "comments1.xml");
        Assert.Contains("tc=", comments, StringComparison.Ordinal);
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
}
