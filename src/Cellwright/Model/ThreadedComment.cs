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
    /// <summary>Extension markup carried through for providers the model does not understand.</summary>
    internal string? ExtensionListXml { get; init; }

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

    /// <summary>Extension markup attached to the person list itself.</summary>
    internal string? ExtensionListXml { get; set; }

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

    /// <summary>Whether this exact person object came from this collection.</summary>
    internal bool Owns(Person person)
    {
        foreach (Person existing in _people)
        {
            if (ReferenceEquals(existing, person))
                return true;
        }

        return false;
    }

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
    private ThreadedCommentCollection? _owner;
    private CellRef _cell;
    private string _text = null!;
    private bool _resolved;

    /// <summary>The original mentions element, including ranges and person identifiers.</summary>
    internal string? MentionsXml { get; set; }

    /// <summary>Extension markup attached to this message.</summary>
    internal string? ExtensionListXml { get; set; }

    /// <summary>Cell the thread is anchored to.</summary>
    public required CellRef Cell
    {
        get => _cell;
        set
        {
            if (_cell == value)
                return;

            _cell = value;
            _owner?.MarkModified();
        }
    }

    /// <summary>What this message says.</summary>
    public required string Text
    {
        get => _text;
        set
        {
            if (string.Equals(_text, value, StringComparison.Ordinal))
                return;

            _text = value;
            bool droppedMentions = MentionsXml is not null;
            bool droppedExtensions = ExtensionListXml is not null;
            MentionsXml = null;
            ExtensionListXml = null;
            _owner?.MarkModified();
            if (droppedMentions || droppedExtensions)
                _owner?.ReportTextMetadataDropped(this, droppedMentions, droppedExtensions);
        }
    }

    /// <summary>Identifier of the person who wrote it.</summary>
    public required string PersonId { get; init; }

    /// <summary>Identifier of this message, which its replies name.</summary>
    public string Id { get; init; } = Person.NewId();

    /// <summary>Identifier of the message this one answers, or null when it starts a thread.</summary>
    public string? ParentId { get; init; }

    /// <summary>When it was written, in UTC.</summary>
    public DateTimeOffset? Created { get; init; }

    /// <summary>Whether this message has been marked resolved.</summary>
    public bool Resolved
    {
        get => _resolved;
        set
        {
            if (_resolved == value)
                return;

            _resolved = value;
            _owner?.MarkModified();
        }
    }

    /// <summary>The replies to this message, in the order they were written.</summary>
    public IReadOnlyList<ThreadedComment> Replies => _replies;

    /// <summary>Whether this message starts a thread rather than answering one.</summary>
    public bool IsFirst => ParentId is null;

    /// <summary>
    /// Whether this loaded message carries mentions or extension markup that Cellwright keeps
    /// opaquely rather than interpreting.
    /// </summary>
    public bool HasOpaqueMetadata => MentionsXml is not null || ExtensionListXml is not null;

    /// <summary>
    /// Discards preserved mentions and message extension markup explicitly. Use this before a
    /// lower-level rewrite when those ranges or checksums are no longer valid. Changing
    /// <see cref="Text"/> performs the same safe drop automatically and emits a warning.
    /// </summary>
    /// <returns>Whether anything was cleared.</returns>
    public bool ClearOpaqueMetadata()
    {
        bool cleared = HasOpaqueMetadata;
        MentionsXml = null;
        ExtensionListXml = null;
        if (cleared)
            _owner?.MarkModified();
        return cleared;
    }

    /// <summary>Attaches a reply, which the collection does while building the tree.</summary>
    internal void AddReply(ThreadedComment reply) => _replies.Add(reply);

    /// <summary>Clears the projected reply tree before it is rebuilt from the flat list.</summary>
    internal void ClearReplies() => _replies.Clear();

    /// <summary>Attaches this message to the collection that owns its serialized state.</summary>
    internal void Attach(ThreadedCommentCollection? owner) => _owner = owner;

    /// <summary>Whether this exact message belongs to a collection.</summary>
    internal bool IsOwnedBy(ThreadedCommentCollection owner) => ReferenceEquals(_owner, owner);

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
    private readonly Worksheet? _worksheet;

    /// <summary>Creates a detached collection. Worksheet collections are created by <see cref="Worksheet"/>.</summary>
    public ThreadedCommentCollection()
    {
    }

    /// <summary>Creates the collection owned by a worksheet.</summary>
    internal ThreadedCommentCollection(Worksheet worksheet) => _worksheet = worksheet;

    /// <inheritdoc/>
    public int Count => _comments.Count;

    /// <inheritdoc/>
    public ThreadedComment this[int index] => _comments[index];

    /// <summary>Whether the model, rather than the loaded file, owns the comments.</summary>
    internal bool IsModified { get; private set; }

    /// <summary>Extension markup attached to the threaded-comments list itself.</summary>
    internal string? ExtensionListXml { get; set; }

    /// <summary>
    /// The first message of every thread, each carrying its replies. A reply whose parent is not
    /// in the sheet is treated as starting a thread of its own rather than dropped.
    /// </summary>
    public IReadOnlyList<ThreadedComment> Threads
    {
        get
        {
            Projection projection = BuildProjection();

            var roots = new List<ThreadedComment>();
            for (int i = 0; i < _comments.Count; i++)
            {
                ThreadedComment comment = _comments[i];
                comment.ClearReplies();
            }

            for (int i = 0; i < _comments.Count; i++)
            {
                ThreadedComment comment = _comments[i];
                int parent = projection.Parents[i];
                if (parent >= 0 && !projection.InCycle[i])
                    _comments[parent].AddReply(comment);
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
        ValidateAuthor(author);

        var comment = new ThreadedComment
        {
            Cell = cell,
            Text = text,
            PersonId = author.Id,
            Created = DateTimeOffset.UtcNow,
        };

        comment.Attach(this);
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
        if (!parent.IsOwnedBy(this))
            throw new ArgumentException("The parent comment does not belong to this worksheet's collection.", nameof(parent));
        ValidateAuthor(author);

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

        reply.Attach(this);
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
        if (removed)
            comment.Attach(null);

        if (removed && comment.IsFirst)
        {
            foreach (ThreadedComment reply in _comments.Where(
                other => string.Equals(other.ParentId, comment.Id, StringComparison.OrdinalIgnoreCase)).ToArray())
            {
                _comments.Remove(reply);
                reply.Attach(null);
            }
        }

        IsModified |= removed;
        return removed;
    }

    /// <summary>Removes every thread on a cell.</summary>
    /// <param name="cell">The cell to clear.</param>
    public bool RemoveAt(CellRef cell)
    {
        ThreadedComment[] matches = [.. _comments.Where(comment => comment.Cell == cell)];
        foreach (ThreadedComment comment in matches)
        {
            _comments.Remove(comment);
            comment.Attach(null);
        }

        int removed = matches.Length;
        IsModified |= removed > 0;
        return removed > 0;
    }

    /// <summary>Adds a message read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(ThreadedComment comment)
    {
        comment.Attach(this);
        _comments.Add(comment);
    }

    /// <summary>Adds reconciled messages and makes their new owning part authoritative.</summary>
    internal void AddReconciled(IEnumerable<ThreadedComment> comments)
    {
        foreach (ThreadedComment comment in comments)
        {
            comment.Attach(this);
            _comments.Add(comment);
        }

        IsModified = true;
    }

    /// <summary>Detaches a projected conversation during workbook-wide reconciliation.</summary>
    internal void RemoveReconciled(IReadOnlyCollection<ThreadedComment> comments)
    {
        if (comments.Count == 0)
            return;

        var removed = new HashSet<ThreadedComment>(comments);
        for (int i = _comments.Count - 1; i >= 0; i--)
        {
            ThreadedComment comment = _comments[i];
            if (!removed.Contains(comment))
                continue;

            _comments.RemoveAt(i);
            comment.Attach(null);
        }

        IsModified = true;
    }

    /// <summary>Marks the collection as owned by the model, e.g. after a structural edit.</summary>
    internal void MarkModified() => IsModified = true;

    /// <summary>Reports malformed imported links once, while load diagnostics are being built.</summary>
    internal void ReportTopologyWarnings(string sheetName, Action<WorkbookWarning> warn)
    {
        Projection projection = BuildProjection();
        if (projection.DuplicateIds > 0)
        {
            warn(new WorkbookWarning(
                "DuplicateThreadedCommentId",
                $"The threaded comments of sheet '{sheetName}' contain {projection.DuplicateIds} duplicate message " +
                "identifier(s). The first message with each identifier owns replies in the projected conversation."));
        }

        if (projection.MissingParents > 0)
        {
            warn(new WorkbookWarning(
                "MissingThreadedCommentParent",
                $"The threaded comments of sheet '{sheetName}' contain {projection.MissingParents} reply link(s) " +
                "whose parent is absent. Those messages are projected as thread roots."));
        }

        int cyclic = 0;
        foreach (bool inCycle in projection.InCycle)
            cyclic += inCycle ? 1 : 0;
        if (cyclic > 0)
        {
            warn(new WorkbookWarning(
                "ThreadedCommentParentCycle",
                $"The threaded comments of sheet '{sheetName}' contain {cyclic} message(s) in parent cycles. " +
                "Cycle members are projected as thread roots."));
        }
    }

    /// <summary>Emits the edit diagnostic after stale text-relative markup has been discarded.</summary>
    internal void ReportTextMetadataDropped(
        ThreadedComment comment, bool droppedMentions, bool droppedExtensions)
    {
        if (_worksheet is null)
            return;

        string what = droppedMentions && droppedExtensions
            ? "mentions and extension markup"
            : droppedMentions ? "mentions" : "extension markup";
        _worksheet.Workbook.OnWarning?.Invoke(new WorkbookWarning(
            "ThreadedCommentTextMetadataDropped",
            $"Changing threaded-comment text on sheet '{_worksheet.Name}' discarded its opaque {what}; " +
            "their text ranges and checksums cannot be rebased safely.",
            comment.Cell));
    }

    /// <summary>Moves every message to the cell the mapping gives it, dropping the rest.</summary>
    /// <remarks>
    /// The collection stays unmodified when the mapping leaves every anchor alone. Preserving
    /// that copy-on-write state avoids rewriting the part. Loaded mentions and extension markup
    /// are also carried through when a rewrite is necessary.
    /// </remarks>
    internal void Reanchor(Func<CellRef, CellRef?> map)
    {
        if (_comments.Count == 0)
            return;

        bool removed = false;
        for (int i = _comments.Count - 1; i >= 0; i--)
        {
            ThreadedComment comment = _comments[i];
            if (map(comment.Cell) is not { } cell)
            {
                _comments.RemoveAt(i);
                comment.Attach(null);
                removed = true;
                continue;
            }

            comment.Cell = cell;
        }

        IsModified |= removed;
    }

    private void ValidateAuthor(Person author)
    {
        if (_worksheet is null)
            return;
        if (!_worksheet.Workbook.People.Owns(author))
            throw new ArgumentException("The author must belong to this worksheet's workbook.", nameof(author));
    }

    /// <summary>
    /// Resolves parent identifiers without recursion. Duplicate identifiers use the first
    /// message in file order; missing links and every member of a cycle become roots.
    /// </summary>
    private Projection BuildProjection()
    {
        int count = _comments.Count;
        var byId = new Dictionary<string, int>(count, StringComparer.OrdinalIgnoreCase);
        int duplicates = 0;
        for (int i = 0; i < count; i++)
        {
            if (!byId.TryAdd(_comments[i].Id, i))
                duplicates++;
        }

        var parents = new int[count];
        Array.Fill(parents, -1);
        int missingParents = 0;
        for (int i = 0; i < count; i++)
        {
            if (_comments[i].ParentId is not { Length: > 0 } parentId)
                continue;

            if (byId.TryGetValue(parentId, out int parent))
                parents[i] = parent;
            else
                missingParents++;
        }

        var inCycle = new bool[count];
        var state = new byte[count];
        var path = new List<int>();
        for (int start = 0; start < count; start++)
        {
            if (state[start] != 0)
                continue;

            path.Clear();
            int current = start;
            while (current >= 0 && state[current] == 0)
            {
                state[current] = 1;
                path.Add(current);
                current = parents[current];
            }

            if (current >= 0 && state[current] == 1)
            {
                int cycleStart = path.Count - 1;
                while (path[cycleStart] != current)
                    cycleStart--;
                for (int i = cycleStart; i < path.Count; i++)
                    inCycle[path[i]] = true;
            }

            foreach (int comment in path)
                state[comment] = 2;
        }

        return new Projection(parents, inCycle, duplicates, missingParents);
    }

    private readonly record struct Projection(
        int[] Parents, bool[] InCycle, int DuplicateIds, int MissingParents);

    /// <inheritdoc/>
    public IEnumerator<ThreadedComment> GetEnumerator() => _comments.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
