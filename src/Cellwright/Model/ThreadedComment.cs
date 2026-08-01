using System.Collections;

namespace Cellwright;

/// <summary>
/// Someone who wrote a threaded comment. Persons are workbook-wide: a comment names its author
/// by identifier, and the identifier is looked up here.
/// </summary>
/// <param name="Id">Identifier the comments refer to, a GUID in braces.</param>
/// <param name="DisplayName">The name shown beside the comment.</param>
/// <param name="UserId">Identifier the provider issued, when there is one.</param>
/// <param name="ProviderId">Which kind of account the identifier belongs to, e.g. <c>None</c> or <c>AD</c>.</param>
public sealed record Person(string Id, string DisplayName, string? UserId = null, string? ProviderId = null)
{
    /// <summary>A person known only by name, with an identifier made up for them.</summary>
    /// <param name="displayName">The name to show.</param>
    public static Person Named(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName);
        return new Person(NewId(), displayName, UserId: null, ProviderId: "None");
    }

    /// <summary>An identifier in the shape the format uses: a GUID in braces, upper case.</summary>
    internal static string NewId() => $"{{{Guid.NewGuid().ToString("D").ToUpperInvariant()}}}";
}

/// <summary>
/// The people who wrote the workbook's threaded comments.
/// </summary>
public sealed class PersonCollection : IReadOnlyList<Person>
{
    private readonly List<Person> _people = [];

    /// <inheritdoc/>
    public int Count => _people.Count;

    /// <inheritdoc/>
    public Person this[int index] => _people[index];

    /// <summary>Whether the model, rather than the loaded file, owns the list.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>The person with an identifier, or <see langword="null"/>.</summary>
    /// <param name="id">The identifier a comment names.</param>
    public Person? ById(string? id)
    {
        if (id is null)
            return null;

        foreach (Person person in _people)
        {
            if (string.Equals(person.Id, id, StringComparison.OrdinalIgnoreCase))
                return person;
        }

        return null;
    }

    /// <summary>Adds a person, or returns the one already there under the same name.</summary>
    /// <param name="displayName">The name to show beside their comments.</param>
    public Person Add(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName);

        foreach (Person existing in _people)
        {
            if (string.Equals(existing.DisplayName, displayName, StringComparison.Ordinal))
                return existing;
        }

        Person person = Person.Named(displayName);
        _people.Add(person);
        IsModified = true;
        return person;
    }

    /// <summary>Adds a person read from a file without marking the list as edited.</summary>
    internal void AddFromFile(Person person) => _people.Add(person);

    /// <summary>Marks the list as owned by the model.</summary>
    internal void MarkModified() => IsModified = true;

    /// <inheritdoc/>
    public IEnumerator<Person> GetEnumerator() => _people.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>
/// One message of a threaded comment — the modern conversation attached to a cell, as opposed to
/// the classic sticky note that <see cref="CellComment"/> models.
/// </summary>
/// <remarks>
/// A thread is a first message and the replies under it, tied together by identifier rather than
/// by nesting: every message names the cell it belongs to, and a reply names the message it
/// answers. <see cref="Replies"/> presents that as the tree it reads as.
/// </remarks>
public sealed class ThreadedComment
{
    private readonly List<ThreadedComment> _replies = [];

    /// <summary>Cell the thread is anchored to.</summary>
    public required CellRef Cell { get; set; }

    /// <summary>What this message says.</summary>
    public required string Text { get; set; }

    /// <summary>Identifier of the person who wrote it.</summary>
    public required string PersonId { get; init; }

    /// <summary>Identifier of this message, which its replies name.</summary>
    public string Id { get; init; } = Person.NewId();

    /// <summary>Identifier of the message this one answers, or null when it starts a thread.</summary>
    public string? ParentId { get; init; }

    /// <summary>When it was written, in UTC.</summary>
    public DateTimeOffset? Created { get; init; }

    /// <summary>Whether the thread has been marked resolved. Only meaningful on a first message.</summary>
    public bool Resolved { get; set; }

    /// <summary>The replies to this message, in the order they were written.</summary>
    public IReadOnlyList<ThreadedComment> Replies => _replies;

    /// <summary>Whether this message starts a thread rather than answering one.</summary>
    public bool IsFirst => ParentId is null;

    /// <summary>Attaches a reply, which the collection does while building the tree.</summary>
    internal void AddReply(ThreadedComment reply) => _replies.Add(reply);

    /// <inheritdoc/>
    public override string ToString() => $"{Cell}: {Text}";
}

/// <summary>
/// The threaded comments of one worksheet.
/// </summary>
/// <remarks>
/// Held flat, in file order, because that is what the part stores and what a save has to write
/// back. <see cref="Threads"/> is the same messages arranged as conversations.
/// </remarks>
public sealed class ThreadedCommentCollection : IReadOnlyList<ThreadedComment>
{
    private readonly List<ThreadedComment> _comments = [];

    /// <inheritdoc/>
    public int Count => _comments.Count;

    /// <inheritdoc/>
    public ThreadedComment this[int index] => _comments[index];

    /// <summary>Whether the model, rather than the loaded file, owns the comments.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>
    /// The first message of every thread, each carrying its replies. A reply whose parent is not
    /// in the sheet is treated as starting a thread of its own rather than dropped.
    /// </summary>
    public IReadOnlyList<ThreadedComment> Threads
    {
        get
        {
            var byId = new Dictionary<string, ThreadedComment>(StringComparer.OrdinalIgnoreCase);
            foreach (ThreadedComment comment in _comments)
                byId[comment.Id] = comment;

            var roots = new List<ThreadedComment>();
            foreach (ThreadedComment comment in _comments)
            {
                if (comment.ParentId is { } parent && byId.TryGetValue(parent, out ThreadedComment? owner))
                    owner.AddReply(comment);
                else
                    roots.Add(comment);
            }

            return roots;
        }
    }

    /// <summary>The messages anchored to one cell, in file order.</summary>
    /// <param name="cell">The cell to look at.</param>
    public IReadOnlyList<ThreadedComment> At(CellRef cell)
    {
        var found = new List<ThreadedComment>();
        foreach (ThreadedComment comment in _comments)
        {
            if (comment.Cell == cell)
                found.Add(comment);
        }

        return found;
    }

    /// <summary>Starts a thread on a cell.</summary>
    /// <param name="cell">Where to anchor it.</param>
    /// <param name="text">What the first message says.</param>
    /// <param name="author">Who wrote it.</param>
    public ThreadedComment Add(CellRef cell, string text, Person author)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(author);

        var comment = new ThreadedComment
        {
            Cell = cell,
            Text = text,
            PersonId = author.Id,
            Created = DateTimeOffset.UtcNow,
        };

        _comments.Add(comment);
        IsModified = true;
        return comment;
    }

    /// <summary>Answers a message, in the same thread and on the same cell.</summary>
    /// <param name="parent">The message being answered.</param>
    /// <param name="text">What the reply says.</param>
    /// <param name="author">Who wrote it.</param>
    public ThreadedComment Reply(ThreadedComment parent, string text, Person author)
    {
        ArgumentNullException.ThrowIfNull(parent);
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(author);

        // A reply answers the first message of the thread, never another reply: the format has
        // one level, and Excel shows a flat conversation.
        string parentId = parent.ParentId ?? parent.Id;

        var reply = new ThreadedComment
        {
            Cell = parent.Cell,
            Text = text,
            PersonId = author.Id,
            ParentId = parentId,
            Created = DateTimeOffset.UtcNow,
        };

        _comments.Add(reply);
        IsModified = true;
        return reply;
    }

    /// <summary>Removes a message and, when it started a thread, everything answering it.</summary>
    /// <param name="comment">The message to remove.</param>
    public bool Remove(ThreadedComment comment)
    {
        ArgumentNullException.ThrowIfNull(comment);

        bool removed = _comments.Remove(comment);
        if (removed && comment.IsFirst)
            _comments.RemoveAll(other => string.Equals(other.ParentId, comment.Id, StringComparison.OrdinalIgnoreCase));

        IsModified |= removed;
        return removed;
    }

    /// <summary>Removes every thread on a cell.</summary>
    /// <param name="cell">The cell to clear.</param>
    public bool RemoveAt(CellRef cell)
    {
        int removed = _comments.RemoveAll(comment => comment.Cell == cell);
        IsModified |= removed > 0;
        return removed > 0;
    }

    /// <summary>Adds a message read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(ThreadedComment comment) => _comments.Add(comment);

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit.</summary>
    internal void MarkModified() => IsModified = true;

    /// <summary>Moves every message to the cell the mapping gives it, dropping the rest.</summary>
    internal void Reanchor(Func<CellRef, CellRef?> map)
    {
        if (_comments.Count == 0)
            return;

        var kept = new List<ThreadedComment>(_comments.Count);
        foreach (ThreadedComment comment in _comments)
        {
            if (map(comment.Cell) is not { } cell)
                continue;

            comment.Cell = cell;
            kept.Add(comment);
        }

        _comments.Clear();
        _comments.AddRange(kept);
        IsModified = true;
    }

    /// <inheritdoc/>
    public IEnumerator<ThreadedComment> GetEnumerator() => _comments.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
