namespace Cellwright.Xls;

/// <summary>What a VBA project is bound to.</summary>
public enum VbaReferenceKind
{
    /// <summary>An automation type library, the ordinary case: a DLL or an OCX.</summary>
    Library,

    /// <summary>Another VBA project, in a different document.</summary>
    Project,

    /// <summary>The type library behind an ActiveX control placed on a form.</summary>
    Control,
}

/// <summary>
/// One library a VBA project binds to. For anyone looking at a document they were sent, this is
/// the first question worth asking: what the code can reach before reading a line of it.
/// </summary>
/// <remarks>
/// The identifier is a packed string (MS-OVBA 2.1.1.8) of the shape
/// <c>*\G{GUID}#major.minor#lcid#path#description</c>, with the leading marker naming the
/// flavour. The pieces are split out here, and the whole string is kept as
/// <see cref="Identifier"/> because a producer is free to write one this shape does not cover.
/// </remarks>
public sealed class VbaReference
{
    private VbaReference(
        string? name, VbaReferenceKind kind, string identifier, Guid? typeLibrary, string? path, string? description)
    {
        Name = name;
        Kind = kind;
        Identifier = identifier;
        TypeLibrary = typeLibrary;
        Path = path;
        Description = description;
    }

    /// <summary>Name the project refers to the library by, or <see langword="null"/> when it declares none.</summary>
    public string? Name { get; }

    /// <summary>What kind of thing is referenced.</summary>
    public VbaReferenceKind Kind { get; }

    /// <summary>The identifier exactly as the file spells it.</summary>
    public string Identifier { get; }

    /// <summary>Class identifier of the type library, when the reference names one.</summary>
    public Guid? TypeLibrary { get; }

    /// <summary>Path the library was registered at, when the reference carries one.</summary>
    public string? Path { get; }

    /// <summary>Human-readable name of the library, when the reference carries one.</summary>
    public string? Description { get; }

    /// <summary>Splits a library identifier into the pieces its grammar defines.</summary>
    internal static VbaReference Parse(string? name, VbaReferenceKind kind, string identifier)
    {
        // Sections are separated by '#', and the first holds the marker and the class id.
        string[] parts = identifier.Split('#');
        Guid? typeLibrary = null;
        if (parts.Length > 0 && ExtractGuid(parts[0]) is { } found)
            typeLibrary = found;

        // *\G{guid} # version # lcid # path # description
        string? path = parts.Length > 3 ? Empty(parts[3]) : null;
        string? description = parts.Length > 4 ? Empty(parts[4]) : null;
        return new VbaReference(Empty(name), kind, identifier, typeLibrary, path, description);

        static string? Empty(string? value) => string.IsNullOrEmpty(value) ? null : value;
    }

    /// <summary>
    /// Pulls the brace-delimited class identifier out of the first section. A reference to a
    /// twiddled library carries all zeros there by convention (MS-OVBA 2.3.4.2.2.3), which is
    /// not an identifier and is reported as none.
    /// </summary>
    private static Guid? ExtractGuid(string section)
    {
        int open = section.IndexOf('{');
        int close = section.IndexOf('}');
        return open >= 0 && close > open &&
            Guid.TryParse(section.AsSpan(open, close - open + 1), out Guid parsed) &&
            parsed != Guid.Empty
            ? parsed
            : null;
    }

    /// <inheritdoc/>
    public override string ToString() => Description is null ? Name ?? Identifier : $"{Name ?? Identifier} ({Description})";
}
