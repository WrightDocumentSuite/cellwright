namespace Cellwright.Pdf.Layout;

/// <summary>
/// The interactive comments of a sheet, arranged into safe PDF threads and indexed by their
/// printable cell anchor.
/// </summary>
/// <remarks>
/// SpreadsheetML stores threaded comments as a flat, identifier-linked list. Loaded input can be
/// malformed, so this projection breaks cycles iteratively, refuses ambiguous duplicate parents,
/// and never follows a recursive model projection. Composition then takes each thread at most once;
/// repeated print-title cells do not duplicate their annotations on every page.
/// </remarks>
internal sealed class SheetCommentPlan
{
    private const string ResolvedMessage =
        "Resolved in the source Excel workbook; resolver identity and time are unavailable.";

    private readonly PdfExportContext _context;
    private readonly Worksheet _sheet;
    private readonly Dictionary<CellRef, List<PdfCommentThread>> _at = [];
    private readonly Dictionary<string, string?> _authors = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<PdfCommentThread> _threads = [];
    private readonly HashSet<PdfCommentThread> _placed = [];

    private int _danglingParents;
    private int _duplicateIdentifiers;
    private int _cycles;
    private int _crossCellReplies;
    private int _unknownAuthors;
    private int _placedResolvedInformation;

    internal SheetCommentPlan(PdfExportContext context, Worksheet sheet)
    {
        _context = context;
        _sheet = sheet;

        foreach (Person person in sheet.Workbook.People)
        {
            if (!string.IsNullOrWhiteSpace(person.Id) && !_authors.ContainsKey(person.Id))
                _authors.Add(person.Id, person.DisplayName);
        }

        context.CommentLosses.Collect(sheet.Workbook, sheet.ThreadedComments);
        AddClassicNotes();
        AddThreadedComments();
    }

    /// <summary>Adds every comment whose one real cell rectangle is visible on this page.</summary>
    internal void AddAt(
        ComposedPage page,
        CellRef cell,
        double x,
        double y,
        double width,
        double height,
        TagRef? tag)
    {
        if (!_at.TryGetValue(cell, out List<PdfCommentThread>? threads))
            return;

        int stackIndex = 0;
        foreach (PdfCommentThread thread in threads)
        {
            if (!_placed.Add(thread))
                continue;

            page.Items.Add(new CommentItem
            {
                X = x,
                Y = y,
                Width = width,
                Height = height,
                Thread = thread,
                StackIndex = stackIndex++,
                Tag = tag,
            });

            _placedResolvedInformation += thread.ResolvedInformationCount;
        }
    }

    /// <summary>Reports comments that could not be anchored and malformed input that was repaired.</summary>
    internal void Complete()
    {
        int hidden = 0;
        int nonRendered = 0;

        foreach (PdfCommentThread thread in _threads)
        {
            if (_placed.Contains(thread))
                continue;

            bool anchorHidden = _sheet.Row(thread.Anchor.Row).Hidden ||
                _sheet.Column(thread.Anchor.Column).Hidden;
            if (anchorHidden)
                hidden += thread.SourceMessageCount;
            else
                nonRendered += thread.SourceMessageCount;
        }

        if (hidden > 0)
        {
            Say(
                $"{hidden} comment message(s) are anchored to a hidden row or column and were not emitted.",
                "comment-hidden");
        }

        if (nonRendered > 0)
        {
            Say(
                $"{nonRendered} comment message(s) have no visible printed cell rectangle and could not become PDF annotations.",
                "comment-anchors");
        }

        if (_danglingParents > 0)
        {
            Say(
                $"{_danglingParents} threaded comment reply/replies name a missing or ambiguous parent; " +
                "they were emitted as independent threads.",
                "comment-thread-dangling");
        }

        int malformed = _duplicateIdentifiers + _cycles + _crossCellReplies;
        if (malformed > 0)
        {
            Say(
                $"Malformed threaded comments were repaired without inventing data: " +
                $"{_duplicateIdentifiers} duplicate identifier(s), {_cycles} cycle(s), " +
                $"and {_crossCellReplies} cross-cell reply/replies.",
                "comment-thread-malformed");
        }

        if (_unknownAuthors > 0)
        {
            Say(
                $"{_unknownAuthors} threaded comment message(s) name no known workbook person; " +
                "their PDF author was left empty.",
                "comment-authors");
        }

        if (_placedResolvedInformation > 0)
        {
            Say(
                $"{_placedResolvedInformation} resolved message(s) do not record the resolver identity required " +
                "by PDF /State /T; machine-readable Completed state was omitted and a provenance-free " +
                "informational reply was emitted.",
                "comment-resolved-state");
        }
    }

    private void AddClassicNotes()
    {
        foreach (CellComment note in _sheet.Comments)
        {
            string name = $"cellwright-note-{_sheet.Id}-{note.Cell.Row}-{note.Cell.Column}";
            var thread = new PdfCommentThread(
                note.Cell,
                [new PdfCommentMessage(note.Text, note.Author, null, name, "Excel note", -1)],
                sourceMessageCount: 1,
                resolvedInformationCount: 0);
            Add(thread);
        }
    }

    private void AddThreadedComments()
    {
        if (_sheet.ThreadedComments.Count == 0)
            return;

        var nodes = new List<ThreadNode>(_sheet.ThreadedComments.Count);
        var groups = new Dictionary<string, List<ThreadNode>>(StringComparer.OrdinalIgnoreCase);

        for (int index = 0; index < _sheet.ThreadedComments.Count; index++)
        {
            ThreadedComment source = _sheet.ThreadedComments[index];
            var node = new ThreadNode(source, index);
            nodes.Add(node);

            if (string.IsNullOrWhiteSpace(source.Id))
                continue;

            if (!groups.TryGetValue(source.Id, out List<ThreadNode>? group))
            {
                group = [];
                groups.Add(source.Id, group);
            }

            group.Add(node);
        }

        var unique = new Dictionary<string, ThreadNode>(StringComparer.OrdinalIgnoreCase);
        foreach ((string id, List<ThreadNode> group) in groups)
        {
            if (group.Count == 1)
                unique.Add(id, group[0]);
            else
                _duplicateIdentifiers++;
        }

        foreach (ThreadNode node in nodes)
        {
            if (node.Source.ParentId is not { } parentId)
                continue;

            if (!unique.TryGetValue(parentId, out ThreadNode? parent))
            {
                _danglingParents++;
                continue;
            }

            node.Parent = parent;
            parent.Children.Add(node);

            if (node.Source.Cell != parent.Source.Cell)
                _crossCellReplies++;
        }

        BreakCycles(nodes);

        foreach (ThreadNode root in nodes.Where(static node => node.Parent is null).OrderBy(static node => node.Index))
            Add(BuildThread(root, groups));
    }

    private void BreakCycles(List<ThreadNode> nodes)
    {
        foreach (ThreadNode start in nodes)
        {
            if (start.State != 0)
                continue;

            var path = new List<ThreadNode>();
            ThreadNode? current = start;
            while (current is not null && current.State == 0)
            {
                current.State = 1;
                path.Add(current);
                current = current.Parent;
            }

            if (current is { State: 1 })
            {
                int cycleStart = path.IndexOf(current);
                ThreadNode breakAt = path[cycleStart];
                for (int i = cycleStart + 1; i < path.Count; i++)
                {
                    if (path[i].Index < breakAt.Index)
                        breakAt = path[i];
                }

                breakAt.Parent!.Children.Remove(breakAt);
                breakAt.Parent = null;
                _cycles++;
            }

            foreach (ThreadNode node in path)
                node.State = 2;
        }
    }

    private PdfCommentThread BuildThread(
        ThreadNode root,
        Dictionary<string, List<ThreadNode>> groups)
    {
        var messages = new List<PdfCommentMessage>();
        var stack = new Stack<(ThreadNode Node, int ParentIndex, bool Resolution)>();
        stack.Push((root, -1, false));
        int sourceMessages = 0;
        int resolvedInformation = 0;

        while (stack.Count > 0)
        {
            (ThreadNode node, int parentIndex, bool resolution) = stack.Pop();
            if (resolution)
            {
                messages.Add(new PdfCommentMessage(
                    ResolvedMessage,
                    null,
                    null,
                    $"{DurableName(node, groups)}-resolved",
                    "Excel thread status",
                    parentIndex));
                resolvedInformation++;
                continue;
            }

            int messageIndex = messages.Count;
            string name = DurableName(node, groups);
            string? author = node.Source.PersonId is { Length: > 0 } personId
                ? _authors.GetValueOrDefault(personId)
                : null;
            if (string.IsNullOrWhiteSpace(author))
            {
                author = null;
                _unknownAuthors++;
            }

            messages.Add(new PdfCommentMessage(
                node.Source.Text,
                author,
                node.Source.Created,
                name,
                parentIndex < 0 ? "Excel threaded comment" : "Excel threaded reply",
                parentIndex));
            sourceMessages++;

            // A source message can carry done itself, including a nested reply. Keep the status
            // beside that exact annotation after its descendants, without inventing resolver data.
            if (node.Source.Resolved)
                stack.Push((node, messageIndex, true));

            for (int i = node.Children.Count - 1; i >= 0; i--)
                stack.Push((node.Children[i], messageIndex, false));
        }

        return new PdfCommentThread(
            root.Source.Cell,
            messages,
            sourceMessageCount: sourceMessages,
            resolvedInformationCount: resolvedInformation);
    }

    private string DurableName(ThreadNode node, Dictionary<string, List<ThreadNode>> groups)
    {
        string durable = string.IsNullOrWhiteSpace(node.Source.Id)
            ? $"message-{node.Index + 1}"
            : node.Source.Id;

        if (!string.IsNullOrWhiteSpace(node.Source.Id) &&
            groups.TryGetValue(node.Source.Id, out List<ThreadNode>? group) &&
            group.Count > 1)
        {
            durable = $"{durable}-{node.Index + 1}";
        }

        return $"cellwright-thread-{_sheet.Id}-{durable}";
    }

    private void Add(PdfCommentThread thread)
    {
        _threads.Add(thread);
        if (!_at.TryGetValue(thread.Anchor, out List<PdfCommentThread>? atCell))
        {
            atCell = [];
            _at.Add(thread.Anchor, atCell);
        }

        atCell.Add(thread);
    }

    private void Say(string message, string subject) =>
        _context.Diagnostics.Add(PdfExportWarningKind.ContentSkipped, message, subject);

    private sealed class ThreadNode(ThreadedComment source, int index)
    {
        internal ThreadedComment Source { get; } = source;

        internal int Index { get; } = index;

        internal ThreadNode? Parent { get; set; }

        internal List<ThreadNode> Children { get; } = [];

        internal int State { get; set; }
    }
}

/// <summary>One PDF annotation thread in parent-before-child order.</summary>
internal sealed class PdfCommentThread(
    CellRef anchor,
    IReadOnlyList<PdfCommentMessage> messages,
    int sourceMessageCount,
    int resolvedInformationCount)
{
    internal CellRef Anchor { get; } = anchor;

    internal IReadOnlyList<PdfCommentMessage> Messages { get; } = messages;

    internal int SourceMessageCount { get; } = sourceMessageCount;

    internal int ResolvedInformationCount { get; } = resolvedInformationCount;
}

/// <summary>One root or reply, with its direct parent already earlier in the list.</summary>
internal readonly record struct PdfCommentMessage(
    string Text,
    string? Author,
    DateTimeOffset? Created,
    string Name,
    string Subject,
    int ParentIndex);
