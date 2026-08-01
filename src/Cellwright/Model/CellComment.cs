using System.Collections;

namespace Cellwright;

/// <summary>
/// A note attached to a cell — the classic yellow sticky, not a threaded comment. Rich text
/// is flattened to plain text: the model stores what the note says, not how it is painted.
/// </summary>
public sealed class CellComment
{
    /// <summary>Cell the note is anchored to.</summary>
    public required CellRef Cell { get; set; }

    /// <summary>Text of the note.</summary>
    public required string Text { get; init; }

    /// <summary>Author, or <see langword="null"/> when the file recorded none.</summary>
    public string? Author { get; init; }

    /// <inheritdoc/>
    public override string ToString() => Author is null ? $"{Cell}: {Text}" : $"{Cell} ({Author}): {Text}";
}

/// <summary>
/// The notes of one worksheet, keyed by cell. Adding a second note to the same cell replaces
/// the first, as Excel does.
/// </summary>
public sealed class CommentCollection : IReadOnlyCollection<CellComment>
{
    private readonly Dictionary<CellRef, CellComment> _comments = [];

    /// <inheritdoc/>
    public int Count => _comments.Count;

    /// <summary>
    /// Whether the model, rather than the loaded file, owns the notes. Until something is
    /// changed the original comment and drawing parts are copied through byte for byte, which
    /// keeps the position and shape of every note box.
    /// </summary>
    internal bool IsModified { get; private set; }

    /// <summary>The note on a cell, or <see langword="null"/>.</summary>
    public CellComment? this[CellRef cell] => _comments.GetValueOrDefault(cell);

    /// <summary>Attaches a note to a cell, replacing any note already there.</summary>
    public CellComment Add(CellRef cell, string text, string? author = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        var comment = new CellComment { Cell = cell, Text = text, Author = author };
        _comments[cell] = comment;
        IsModified = true;
        return comment;
    }

    /// <summary>Attaches a note to a cell given as an A1 reference.</summary>
    public CellComment Add(string a1, string text, string? author = null) =>
        Add(CellRef.Parse(a1), text, author);

    /// <summary>Removes the note on a cell. Returns whether one was there.</summary>
    public bool Remove(CellRef cell)
    {
        bool removed = _comments.Remove(cell);
        IsModified |= removed;
        return removed;
    }

    /// <summary>Adds a note read from a file without marking the collection as edited.</summary>
    internal void AddFromFile(CellComment comment) => _comments[comment.Cell] = comment;

    /// <summary>
    /// Moves every note to the cell the mapping gives it, dropping the ones it maps to nothing.
    /// Used when rows or columns are inserted or deleted underneath them.
    /// </summary>
    internal void Reanchor(Func<CellRef, CellRef?> map)
    {
        if (_comments.Count == 0)
            return;

        var moved = new Dictionary<CellRef, CellComment>(_comments.Count);
        foreach (CellComment comment in _comments.Values)
        {
            if (map(comment.Cell) is not { } cell)
                continue;

            comment.Cell = cell;
            moved[cell] = comment;
        }

        _comments.Clear();
        foreach ((CellRef cell, CellComment comment) in moved)
            _comments[cell] = comment;
        IsModified = true;
    }

    /// <inheritdoc/>
    public IEnumerator<CellComment> GetEnumerator() => _comments.Values.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
