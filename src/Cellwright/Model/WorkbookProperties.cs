using System.Collections;

namespace Cellwright;

/// <summary>
/// What the file says about itself: the fields Excel shows under File → Info, plus any custom
/// properties beside them. They live in their own parts of the package rather than in the
/// workbook, which is why a workbook created from scratch used to arrive without any.
/// </summary>
/// <remarks>
/// Fields the model does not name — the language, revision and identifier of the core part,
/// the statistics of the extended one — are kept as they were read and written back untouched.
/// Both parts allow their children in any order, so replaying them beside generated ones is
/// safe. The two exceptions are the lists of sheet names, which are rebuilt on save because
/// the model may well have changed them.
/// </remarks>
public sealed class WorkbookProperties
{
    private DateTime? _modified;
    private bool _modifiedIsExplicit;

    /// <summary>Title of the document.</summary>
    public string? Title { get; set; }

    /// <summary>What the document is about.</summary>
    public string? Subject { get; set; }

    /// <summary>Who wrote it (<c>dc:creator</c>).</summary>
    public string? Author { get; set; }

    /// <summary>Search keywords, separated however the producer chose.</summary>
    public string? Keywords { get; set; }

    /// <summary>Free-text description (<c>dc:description</c>, shown as "Comments" in Excel).</summary>
    public string? Description { get; set; }

    /// <summary>Category the document is filed under.</summary>
    public string? Category { get; set; }

    /// <summary>Workflow status, such as "Draft" or "Final".</summary>
    public string? ContentStatus { get; set; }

    /// <summary>Who saved it last.</summary>
    public string? LastModifiedBy { get; set; }

    /// <summary>When the document was created. Stamped on the first save when nothing set it.</summary>
    public DateTime? Created { get; set; }

    /// <summary>
    /// When the document was last saved. Refreshed on every save unless you assign it
    /// yourself, in which case the value you set is written as it stands.
    /// </summary>
    public DateTime? Modified
    {
        get => _modified;
        set
        {
            _modified = value;
            _modifiedIsExplicit = true;
        }
    }

    /// <summary>Program that produced the file. Written as <c>Cellwright</c> when nothing set it.</summary>
    public string? Application { get; set; }

    /// <summary>Company the document belongs to.</summary>
    public string? Company { get; set; }

    /// <summary>Manager named in the extended properties.</summary>
    public string? Manager { get; set; }

    /// <summary>Properties of the caller's own choosing, in the order they were added.</summary>
    public CustomPropertyCollection Custom { get; } = new();

    /// <summary>Children of the core part the model does not name, replayed verbatim on save.</summary>
    internal string? CoreMarkup { get; set; }

    /// <summary>Children of the extended part the model does not name, replayed verbatim on save.</summary>
    internal string? ExtendedMarkup { get; set; }

    /// <summary>Records the timestamp a file carried, without counting as the caller setting one.</summary>
    internal void SetLoadedModified(DateTime? value) => _modified = value;

    /// <summary>Takes over what the reader found, timestamps included but still not counted as set.</summary>
    internal void CopyFrom(WorkbookProperties source)
    {
        Title = source.Title;
        Subject = source.Subject;
        Author = source.Author;
        Keywords = source.Keywords;
        Description = source.Description;
        Category = source.Category;
        ContentStatus = source.ContentStatus;
        LastModifiedBy = source.LastModifiedBy;
        Created = source.Created;
        Application = source.Application;
        Company = source.Company;
        Manager = source.Manager;
        CoreMarkup = source.CoreMarkup;
        ExtendedMarkup = source.ExtendedMarkup;
        SetLoadedModified(source.Modified);

        foreach ((string name, CellValue value) in source.Custom)
            Custom[name] = value;
    }

    /// <summary>The timestamp to write: the caller's if they chose one, otherwise the moment of saving.</summary>
    internal DateTime ModifiedOnSave() =>
        _modifiedIsExplicit && _modified is { } explicitValue ? explicitValue : DateTime.UtcNow;
}

/// <summary>
/// Custom document properties: named values of the few types the format allows — text, number,
/// yes-or-no and date — kept in the order they were declared.
/// </summary>
public sealed class CustomPropertyCollection : IReadOnlyCollection<KeyValuePair<string, CellValue>>
{
    private readonly OrderedDictionary<string, CellValue> _values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>How many custom properties there are.</summary>
    public int Count => _values.Count;

    /// <summary>
    /// Reads or writes one property by name, matched the way Excel matches them: without
    /// regard for case. Reading an absent name gives <see cref="CellValue.Blank"/>.
    /// </summary>
    public CellValue this[string name]
    {
        get => _values.TryGetValue(name, out CellValue value) ? value : CellValue.Blank;
        set
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            _values[name] = value;
        }
    }

    /// <summary>Whether a property of that name exists.</summary>
    public bool Contains(string name) => _values.ContainsKey(name);

    /// <summary>Removes a property. Returns whether there was one.</summary>
    public bool Remove(string name) => _values.Remove(name);

    /// <summary>Removes every custom property.</summary>
    public void Clear() => _values.Clear();

    /// <inheritdoc/>
    public IEnumerator<KeyValuePair<string, CellValue>> GetEnumerator() => _values.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
