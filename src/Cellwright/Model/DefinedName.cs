using System.Collections;

namespace Cellwright;

/// <summary>
/// A named reference or constant, e.g. <c>TaxRate</c> pointing at <c>Rates!$B$2</c>.
/// Names are workbook-scoped unless a <see cref="Scope"/> sheet is given.
/// </summary>
public sealed class DefinedName
{
    internal DefinedName(string name, string refersTo, Worksheet? scope)
    {
        Name = name;
        RefersTo = refersTo;
        Scope = scope;
    }

    /// <summary>The name as written in formulas.</summary>
    public string Name { get; }

    /// <summary>
    /// Whether this is one of Excel's built-in names (<c>_xlnm.Print_Area</c>,
    /// <c>_xlnm._FilterDatabase</c> and friends), which record sheet settings rather than a
    /// reference somebody named. They round-trip but are not meant to be used in formulas.
    /// </summary>
    public bool IsBuiltIn => Name.StartsWith("_xlnm.", StringComparison.Ordinal);

    /// <summary>Whether Excel hides the name from the name manager.</summary>
    public bool IsHidden { get; set; }

    /// <summary>
    /// What the name resolves to: a formula without the leading <c>=</c>,
    /// usually an absolute reference such as <c>Sheet1!$A$1:$B$5</c>.
    /// </summary>
    public string RefersTo { get; internal set; }

    /// <summary>Sheet the name is local to, or <see langword="null"/> for workbook scope.</summary>
    public Worksheet? Scope { get; }

    /// <summary>Attributes of the loaded element the model does not represent, written back as read.</summary>
    internal List<KeyValuePair<string, string>>? Markup { get; set; }

    /// <inheritdoc/>
    public override string ToString() => Scope is null ? $"{Name} = {RefersTo}" : $"{Scope.Name}!{Name} = {RefersTo}";
}

/// <summary>
/// The defined names of a <see cref="Workbook"/>. Lookup follows Excel's rules:
/// a sheet-scoped name shadows a workbook-scoped one of the same name.
/// </summary>
public sealed class DefinedNameCollection : IReadOnlyCollection<DefinedName>
{
    private readonly List<DefinedName> _names = [];

    /// <inheritdoc/>
    public int Count => _names.Count;

    /// <summary>Raised when a name is added, removed or pointed somewhere else.</summary>
    internal Action? Changed { get; set; }

    /// <summary>Adds or replaces a name.</summary>
    /// <param name="name">Name as used in formulas.</param>
    /// <param name="refersTo">Target formula, e.g. <c>Sheet1!$A$1</c>.</param>
    /// <param name="scope">Sheet the name is local to, or <see langword="null"/> for workbook scope.</param>
    public DefinedName Add(string name, string refersTo, Worksheet? scope = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrEmpty(refersTo);
        if (refersTo.StartsWith('='))
            refersTo = refersTo[1..];

        if (Find(name, scope) is { } existing)
        {
            existing.RefersTo = refersTo;
            Changed?.Invoke();
            return existing;
        }

        var defined = new DefinedName(name, refersTo, scope);
        _names.Add(defined);
        Changed?.Invoke();
        return defined;
    }

    /// <summary>
    /// Resolves a name visible from <paramref name="scope"/>: the sheet-scoped
    /// definition wins, then the workbook-scoped one.
    /// </summary>
    public DefinedName? Resolve(string name, Worksheet? scope)
    {
        DefinedName? global = null;
        foreach (DefinedName defined in _names)
        {
            if (!string.Equals(defined.Name, name, StringComparison.OrdinalIgnoreCase))
                continue;
            if (ReferenceEquals(defined.Scope, scope) && scope is not null)
                return defined;
            if (defined.Scope is null)
                global = defined;
        }

        return global;
    }

    /// <summary>Finds an exact definition by name and scope.</summary>
    public DefinedName? Find(string name, Worksheet? scope = null)
    {
        foreach (DefinedName defined in _names)
        {
            if (string.Equals(defined.Name, name, StringComparison.OrdinalIgnoreCase) &&
                ReferenceEquals(defined.Scope, scope))
            {
                return defined;
            }
        }

        return null;
    }

    /// <summary>Removes a definition. Returns whether anything was removed.</summary>
    public bool Remove(string name, Worksheet? scope = null)
    {
        if (Find(name, scope) is not { } defined || !_names.Remove(defined))
            return false;

        Changed?.Invoke();
        return true;
    }

    /// <inheritdoc/>
    public IEnumerator<DefinedName> GetEnumerator() => _names.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
