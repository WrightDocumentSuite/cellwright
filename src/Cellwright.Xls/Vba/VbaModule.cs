namespace Cellwright.Xls;

/// <summary>What kind of module a piece of VBA lives in.</summary>
public enum VbaModuleKind
{
    /// <summary>A plain module: the loose subs and functions of a project.</summary>
    Procedural,

    /// <summary>The code behind a worksheet or the workbook itself.</summary>
    Document,

    /// <summary>A class module.</summary>
    Class,

    /// <summary>The code behind a user form.</summary>
    Designer,
}

/// <summary>
/// One module of a VBA project, with its source as the editor would show it.
/// </summary>
public sealed class VbaModule
{
    internal VbaModule(
        string name,
        string streamName,
        VbaModuleKind kind,
        string code,
        bool isReadOnly,
        bool isPrivate,
        string? description)
    {
        Name = name;
        StreamName = streamName;
        Kind = kind;
        Code = code;
        IsReadOnly = isReadOnly;
        IsPrivate = isPrivate;
        Description = description;
    }

    /// <summary>Name of the module as the editor shows it.</summary>
    public string Name { get; }

    /// <summary>Name of the stream the module was read from; usually the same as <see cref="Name"/>.</summary>
    public string StreamName { get; }

    /// <summary>What kind of module it is.</summary>
    public VbaModuleKind Kind { get; }

    /// <summary>
    /// The source, lines separated by carriage return and line feed as VBA stores them. The
    /// attribute lines the editor hides — <c>Attribute VB_Name</c> and its relatives — are part
    /// of the text, because they are part of what the file holds.
    /// </summary>
    public string Code { get; }

    /// <summary>Whether the editor refuses to change the module.</summary>
    public bool IsReadOnly { get; }

    /// <summary>Whether the module is hidden from other projects.</summary>
    public bool IsPrivate { get; }

    /// <summary>Description the editor shows beside the module, or <see langword="null"/>.</summary>
    public string? Description { get; }

    /// <inheritdoc/>
    public override string ToString() => $"{Kind} {Name} ({Code.Length} chars)";
}
