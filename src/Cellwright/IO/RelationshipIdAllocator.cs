namespace Cellwright.IO;

/// <summary>
/// Hands out relationship ids that do not collide with the ones a loaded package already
/// uses. Preserved markup points at its relationships by id — a pivot cache, an external
/// reference or a drawing would attach to the wrong target if we reused rId1 blindly.
/// </summary>
internal sealed class RelationshipIdAllocator
{
    private readonly HashSet<string> _used = new(StringComparer.OrdinalIgnoreCase);
    private int _next = 1;

    /// <summary>Marks an existing id as taken.</summary>
    public void Reserve(string id) => _used.Add(id);

    /// <summary>Marks the ids of existing relationships as taken.</summary>
    public void Reserve(IEnumerable<OpcRelationship> relationships)
    {
        foreach (OpcRelationship relationship in relationships)
            _used.Add(relationship.Id);
    }

    /// <summary>Returns an unused id of the conventional <c>rIdN</c> shape.</summary>
    public string Next()
    {
        string id;
        do
        {
            id = "rId" + _next.ToString(System.Globalization.CultureInfo.InvariantCulture);
            _next++;
        }
        while (!_used.Add(id));

        return id;
    }
}
