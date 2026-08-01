using System.Text;
using Cellwright.IO;
using Cellwright.Xls.Cfb;
using Cellwright.Xls.Vba;

namespace Cellwright.Xls;

/// <summary>
/// The macros of a workbook, read out as source. Both formats keep a VBA project in the same
/// shape — a compound file of its own, whose module streams hold compressed source behind a
/// cache the editor uses — so the same reader serves an <c>.xlsm</c> and an <c>.xls</c>.
/// </summary>
/// <remarks>
/// <para>
/// Reading only. A project travels through a save untouched anyway, because the part that holds
/// it is one of those the model carries verbatim, so there is nothing to write back.
/// </para>
/// <para>
/// Locking a project in the editor does not encrypt it: the password stops the editor from
/// showing the code, not a reader from finding it. Source comes out of a locked project exactly
/// as it does out of an open one, which is the whole point for anyone inspecting a file they
/// were sent.
/// </para>
/// </remarks>
/// <example>
/// <code>
/// VbaProject? macros = await VbaProject.LoadAsync("отчёт.xlsm");
/// foreach (VbaModule module in macros?.Modules ?? [])
///     Console.WriteLine($"{module.Kind} {module.Name}\n{module.Code}");
/// </code>
/// </example>
public sealed class VbaProject
{
    /// <summary>Storage an .xls keeps its project in, beside the workbook stream.</summary>
    private const string LegacyStorage = "_VBA_PROJECT_CUR";

    /// <summary>Storage holding the directory and the module streams.</summary>
    private const string ModuleStorage = "VBA";

    /// <summary>Stream naming the modules and recording the project's protection.</summary>
    private const string ProjectStream = "PROJECT";

    private VbaProject(
        string? name,
        string? description,
        string? constants,
        int codePage,
        bool isPasswordProtected,
        VbaProtectionState protection,
        IReadOnlyList<VbaReference> references,
        IReadOnlyList<VbaModule> modules)
    {
        Name = name;
        Description = description;
        CompilationConstants = constants;
        CodePage = codePage;
        IsPasswordProtected = isPasswordProtected;
        Protection = protection;
        References = references;
        Modules = modules;
    }

    /// <summary>Name of the project, as the editor's tree shows it; usually <c>VBAProject</c>.</summary>
    public string? Name { get; }

    /// <summary>Description of the project, or <see langword="null"/> when it carries none.</summary>
    public string? Description { get; }

    /// <summary>
    /// Conditional compilation constants, in the <c>Name = Value : Name = Value</c> form the
    /// editor's project properties dialog uses, or <see langword="null"/> when there are none.
    /// </summary>
    public string? CompilationConstants { get; }

    /// <summary>Code page the project's names and source were written in.</summary>
    public int CodePage { get; }

    /// <summary>
    /// Whether the project carries a password. The source is still readable here: the password
    /// guards the editor, not the file.
    /// </summary>
    public bool IsPasswordProtected { get; }

    /// <summary>Who closed the project, and whether the editor shows it at all.</summary>
    public VbaProtectionState Protection { get; }

    /// <summary>
    /// Libraries the project binds to. What the code can reach, before reading a line of it.
    /// </summary>
    public IReadOnlyList<VbaReference> References { get; }

    /// <summary>The modules, in the order the project declares them.</summary>
    public IReadOnlyList<VbaModule> Modules { get; }

    /// <summary>Finds a module by name, the way VBA compares them: without regard for case.</summary>
    public VbaModule? Find(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        foreach (VbaModule module in Modules)
        {
            if (string.Equals(module.Name, name, StringComparison.OrdinalIgnoreCase))
                return module;
        }

        return null;
    }

    /// <summary>Reads the macros of a file, or returns <see langword="null"/> when it has none.</summary>
    /// <param name="path">Path of an .xlsm, .xlsx, .xls, or a bare <c>vbaProject.bin</c>.</param>
    /// <param name="options">Reading settings; only the fallback code page matters here.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The file holds a project that cannot be unpacked.</exception>
    public static async ValueTask<VbaProject?> LoadAsync(
        string path, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        await using FileStream stream = new(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 16, useAsync: true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads the macros of a stream, or returns <see langword="null"/> when it has none.</summary>
    /// <param name="stream">Stream positioned at the start of the file.</param>
    /// <param name="options">Reading settings; only the fallback code page matters here.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The stream holds a project that cannot be unpacked.</exception>
    public static async ValueTask<VbaProject?> LoadAsync(
        Stream stream, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        options ??= new XlsLoadOptions();

        // A package is read through the zip, one part of it: the project is a few kilobytes
        // even in a workbook of hundreds of megabytes. Only a compound file has to be held
        // whole, because its sectors are scattered and the directory indexes them by offset.
        if (await LooksLikePackageAsync(stream, cancellationToken).ConfigureAwait(false))
        {
            byte[]? found = await ExtractPartAsync(stream, options.LeaveOpen, options.Limits, cancellationToken).ConfigureAwait(false);
            return found is null || !CompoundFile.HasSignature(found)
                ? null
                : Read(CompoundFile.Parse(found, options.Limits), Fallback(options));
        }

        byte[] bytes = await XlsWorkbook.ReadAllAsync(stream, options.LeaveOpen, options.Limits, cancellationToken).ConfigureAwait(false);
        if (CompoundFile.HasSignature(bytes))
            return Read(CompoundFile.Parse(bytes, options.Limits), Fallback(options));

        // Not seekable, so the signature could not be read ahead: the bytes are in hand now.
        using var package = new MemoryStream(bytes, writable: false);
        byte[]? part = await ExtractPartAsync(package, leaveOpen: true, options.Limits, cancellationToken).ConfigureAwait(false);
        return part is null || !CompoundFile.HasSignature(part)
            ? null
            : Read(CompoundFile.Parse(part, options.Limits), Fallback(options));
    }

    /// <summary>
    /// Whether the stream opens with the zip signature. A stream that cannot seek gives no
    /// answer: it has to be read to be recognized, and reading it is what the caller pays for.
    /// </summary>
    private static async ValueTask<bool> LooksLikePackageAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (!stream.CanSeek)
            return false;

        long origin = stream.Position;
        byte[] head = new byte[2];
        int read = await stream.ReadAtLeastAsync(head, head.Length, throwOnEndOfStream: false, cancellationToken)
            .ConfigureAwait(false);
        stream.Position = origin;
        return read == 2 && head[0] == 0x50 && head[1] == 0x4B;
    }

    /// <summary>
    /// Reads the macros carried by a loaded workbook. Available only while the workbook keeps
    /// the parts the model does not represent, which is the default.
    /// </summary>
    internal static VbaProject? FromWorkbook(Workbook workbook, XlsLoadOptions? options)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        if (workbook.Preserved is not { } preserved)
            return null;

        byte[]? part = FindProjectPart(preserved);
        XlsLoadOptions effective = options ?? new XlsLoadOptions();
        return part is null || !CompoundFile.HasSignature(part)
            ? null
            : Read(CompoundFile.Parse(part, effective.Limits), Fallback(effective));
    }

    private static Encoding Fallback(XlsLoadOptions options) =>
        VbaDirectory.Resolve(options.FallbackCodePage, Encoding.Latin1);

    /// <summary>
    /// Locates the project part of a preserved package: by the relationship Office declares for
    /// it, and failing that by the name every producer uses anyway.
    /// </summary>
    private static byte[]? FindProjectPart(PreservedPackage preserved)
    {
        const string relationship = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        foreach (OpcRelationship candidate in preserved.RelationshipsOf(preserved.WorkbookPartPath))
        {
            if (candidate.Type != relationship || candidate.IsExternal)
                continue;

            string target = OpcPath.Resolve(preserved.WorkbookPartPath, candidate.Target);
            if (preserved.Parts.TryGetValue(target, out byte[]? found))
                return found;
        }

        return preserved.Parts.TryGetValue("/xl/vbaProject.bin", out byte[]? canonical) ? canonical : null;
    }

    /// <summary>Pulls the project part out of a package without building a workbook from it.</summary>
    private static async ValueTask<byte[]?> ExtractPartAsync(
        Stream package, bool leaveOpen, PackageLimits limits, CancellationToken cancellationToken)
    {
        OpcPackage opc;
        try
        {
            opc = await OpcPackage.OpenReadAsync(package, leaveOpen, cancellationToken, limits).ConfigureAwait(false);
        }
        catch (InvalidDataException)
        {
            // Neither a compound file nor a zip: nothing here holds a VBA project.
            return null;
        }

        await using (opc.ConfigureAwait(false))
        {
            foreach (string partPath in opc.PartPaths)
            {
                if (partPath.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase))
                    return await opc.ReadPartBytesAsync(partPath, cancellationToken).ConfigureAwait(false);
            }
        }

        return null;
    }

    /// <summary>
    /// Reads a project out of the container holding it. An .xls keeps one beside its workbook
    /// stream under a storage of its own; a package keeps the same thing as a whole part.
    /// </summary>
    private static VbaProject? Read(CompoundFile container, Encoding fallback)
    {
        string[] root =
            container.StorageExists(LegacyStorage) ? [LegacyStorage] :
            container.StorageExists(ModuleStorage) ? [] :
            null!;
        if (root is null)
            return null;

        if (container.TryReadPath([.. root, ModuleStorage, "dir"]) is not { Length: > 0 } packed)
            return null;

        VbaDirectory.Contents directory = VbaDirectory.Parse(VbaCompression.Decompress(packed), fallback);
        Encoding text = VbaDirectory.Resolve(directory.CodePage, fallback);
        ProjectStreamContents declared = ReadProjectStream(container, root, text);

        List<VbaModule> modules = [];
        foreach (VbaModuleEntry entry in directory.Modules)
        {
            if (container.TryReadPath([.. root, ModuleStorage, entry.StreamName]) is not { } stream ||
                entry.SourceOffset >= stream.Length)
            {
                continue;
            }

            // Everything before the offset is a cache the editor keeps for itself; the source
            // is the compressed container that follows it.
            string code = text.GetString(VbaCompression.Decompress(stream.AsSpan(entry.SourceOffset)));
            VbaModuleKind kind = entry.IsProcedural
                ? VbaModuleKind.Procedural
                : declared.Kinds.GetValueOrDefault(entry.Name, VbaModuleKind.Document);
            modules.Add(new VbaModule(
                entry.Name, entry.StreamName, kind, code, entry.IsReadOnly, entry.IsPrivate, entry.Description));
        }

        return new VbaProject(
            directory.Name ?? declared.Name,
            directory.Description,
            directory.Constants,
            directory.CodePage,
            declared.IsProtected,
            declared.Protection,
            directory.References,
            modules);
    }

    /// <summary>What the <c>PROJECT</c> stream adds: module kinds, the password and the lock.</summary>
    private readonly record struct ProjectStreamContents(
        string? Name,
        bool IsProtected,
        VbaProtectionState Protection,
        IReadOnlyDictionary<string, VbaModuleKind> Kinds);

    /// <summary>
    /// Reads the <c>PROJECT</c> stream, a plain list of <c>key=value</c> lines. The directory
    /// lumps classes, forms and document modules into one type, and this is what tells them
    /// apart. A missing or unreadable stream costs only that distinction.
    /// </summary>
    private static ProjectStreamContents ReadProjectStream(CompoundFile container, string[] root, Encoding text)
    {
        var kinds = new Dictionary<string, VbaModuleKind>(StringComparer.OrdinalIgnoreCase);
        if (container.TryReadPath([.. root, ProjectStream]) is not { Length: > 0 } bytes)
            return new ProjectStreamContents(null, false, VbaProtectionState.Open, kinds);

        string? name = null;
        bool protectedProject = false;
        (bool user, bool host, bool editor) locks = default;
        bool visible = true;
        foreach (string line in text.GetString(bytes).Split('\n'))
        {
            ReadOnlySpan<char> trimmed = line.AsSpan().Trim();
            int separator = trimmed.IndexOf('=');
            if (separator <= 0)
                continue;

            ReadOnlySpan<char> key = trimmed[..separator];
            ReadOnlySpan<char> value = trimmed[(separator + 1)..].Trim('"');
            switch (key)
            {
                // A document entry names the sheet and its cookie: "Sheet1/&H00000000".
                case "Document":
                    Declare(kinds, value[..IndexOrEnd(value, '/')], VbaModuleKind.Document);
                    break;
                case "Module":
                    Declare(kinds, value, VbaModuleKind.Procedural);
                    break;
                case "Class":
                    Declare(kinds, value, VbaModuleKind.Class);
                    break;
                case "BaseClass":
                    Declare(kinds, value, VbaModuleKind.Designer);
                    break;
                case "Name":
                    name = value.ToString();
                    break;

                // Every project Excel saves carries these three, protected or not, so what they
                // decode to is the only thing that answers the question.
                case "DPB":
                    protectedProject = VbaProtection.IsPasswordProtected(value);
                    break;
                case "CMG":
                    locks = VbaProtection.DecodeProtectionState(value);
                    break;
                case "GC":
                    visible = VbaProtection.DecodeVisibility(value);
                    break;
            }
        }

        var protection = new VbaProtectionState(locks.user, locks.host, locks.editor, visible);
        return new ProjectStreamContents(name, protectedProject, protection, kinds);

        static void Declare(Dictionary<string, VbaModuleKind> into, ReadOnlySpan<char> module, VbaModuleKind kind)
        {
            if (!module.IsEmpty)
                into[module.ToString()] = kind;
        }

        static int IndexOrEnd(ReadOnlySpan<char> value, char separator) =>
            value.IndexOf(separator) is int found && found >= 0 ? found : value.Length;
    }
}
