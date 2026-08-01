using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Vba;

/// <summary>
/// One module as the project directory describes it: where its stream is and how far into that
/// stream the source begins.
/// </summary>
/// <param name="Name">Name the module has in the editor.</param>
/// <param name="StreamName">Name of the stream inside the <c>VBA</c> storage holding it.</param>
/// <param name="SourceOffset">Bytes of performance cache before the compressed source.</param>
/// <param name="IsProcedural">Whether it is a plain module rather than a class, form or document.</param>
/// <param name="IsReadOnly">Whether the editor refuses to change it.</param>
/// <param name="IsPrivate">Whether it is hidden from other projects.</param>
/// <param name="Description">Description the editor shows, or <see langword="null"/>.</param>
internal readonly record struct VbaModuleEntry(
    string Name,
    string StreamName,
    int SourceOffset,
    bool IsProcedural,
    bool IsReadOnly,
    bool IsPrivate,
    string? Description);

/// <summary>
/// The <c>dir</c> stream of a VBA project (MS-OVBA 2.3.4.2): a run of records describing the
/// project and every module in it. The stream is compressed, so it arrives here already
/// unpacked.
/// </summary>
/// <remarks>
/// Every record is an identifier, a size and that many bytes of payload — with one exception.
/// <c>PROJECTVERSION</c> declares a size of four and then writes six bytes, and a reader that
/// believes it loses its place for the rest of the stream. Everything else, including the
/// reference records whose shapes look irregular in the specification, sizes itself honestly.
/// </remarks>
internal static class VbaDirectory
{
    private const ushort ProjectCodePage = 0x0003;
    private const ushort ProjectName = 0x0004;
    private const ushort ProjectDocString = 0x0005;
    private const ushort ProjectVersion = 0x0009;
    private const ushort ProjectConstants = 0x000C;
    private const ushort ProjectCompatVersion = 0x004A;
    private const ushort ModuleName = 0x0019;
    private const ushort ModuleNameUnicode = 0x0047;
    private const ushort ModuleStreamName = 0x001A;
    private const ushort ModuleStreamNameUnicode = 0x0032;
    private const ushort ModuleDocString = 0x001C;
    private const ushort ModuleOffset = 0x0031;
    private const ushort ModuleTypeProcedural = 0x0021;
    private const ushort ModuleTypeOther = 0x0022;
    private const ushort ModuleReadOnly = 0x0025;
    private const ushort ModulePrivate = 0x0028;
    private const ushort ModuleTerminator = 0x002B;

    // The three shapes a reference takes, plus the name record that may precede any of them.
    private const ushort ReferenceName = 0x0016;
    private const ushort ReferenceRegistered = 0x000D;
    private const ushort ReferenceProject = 0x000E;
    private const ushort ReferenceControl = 0x002F;

    /// <summary>What the directory says about the project as a whole, its references and its modules.</summary>
    /// <param name="CodePage">Code page the project's byte strings are written in.</param>
    /// <param name="Name">Name of the project, or <see langword="null"/> when it declares none.</param>
    /// <param name="Description">Description of the project, or <see langword="null"/>.</param>
    /// <param name="Constants">Conditional compilation constants, or <see langword="null"/>.</param>
    /// <param name="References">Libraries the project binds to, in declaration order.</param>
    /// <param name="Modules">Modules in declaration order.</param>
    public readonly record struct Contents(
        int CodePage,
        string? Name,
        string? Description,
        string? Constants,
        IReadOnlyList<VbaReference> References,
        IReadOnlyList<VbaModuleEntry> Modules);

    /// <summary>Reads the records of an unpacked <c>dir</c> stream.</summary>
    public static Contents Parse(ReadOnlySpan<byte> dir, Encoding fallback)
    {
        Encoding text = fallback;
        int codePage = 0;
        string? projectName = null;
        string? description = null;
        string? constants = null;
        List<VbaModuleEntry> modules = [];
        List<VbaReference> references = [];
        string? pendingReferenceName = null;
        var current = new ModuleBuilder();

        int position = 0;
        while (position + 6 <= dir.Length)
        {
            ushort id = BinaryPrimitives.ReadUInt16LittleEndian(dir[position..]);
            int size = (int)BinaryPrimitives.ReadUInt32LittleEndian(dir[(position + 2)..]);

            // The one record that misstates its own length; six bytes follow, not four.
            if (id == ProjectVersion)
                size = 6;

            if (size < 0 || position + 6 + size > dir.Length)
                break;

            ReadOnlySpan<byte> body = dir.Slice(position + 6, size);
            position += 6 + size;

            switch (id)
            {
                case ProjectCodePage when size >= 2:
                    codePage = BinaryPrimitives.ReadUInt16LittleEndian(body);
                    text = Resolve(codePage, fallback);
                    break;
                case ProjectName:
                    projectName = text.GetString(body);
                    break;
                case ProjectDocString:
                    description = Trimmed(text.GetString(body));
                    break;
                case ProjectConstants:
                    constants = Trimmed(text.GetString(body));
                    break;

                // A name record belongs to the reference that follows it, and is optional.
                case ReferenceName:
                    pendingReferenceName = text.GetString(body);
                    break;
                case ReferenceRegistered or ReferenceProject or ReferenceControl:
                    references.Add(ReadReference(id, body, text, ref pendingReferenceName));
                    break;

                // A module name opens a module; everything until the terminator belongs to it.
                case ModuleName:
                    current.Finish(modules);
                    current.Name = text.GetString(body);
                    break;
                case ModuleDocString:
                    current.Description = Trimmed(text.GetString(body));
                    break;
                case ModuleNameUnicode:
                    current.Name = Encoding.Unicode.GetString(body);
                    break;
                case ModuleStreamName:
                    current.StreamName = text.GetString(body);
                    break;
                case ModuleStreamNameUnicode:
                    current.StreamName = Encoding.Unicode.GetString(body);
                    break;
                case ModuleOffset when size >= 4:
                    current.SourceOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
                    break;
                case ModuleTypeProcedural:
                    current.IsProcedural = true;
                    break;
                case ModuleTypeOther:
                    current.IsProcedural = false;
                    break;
                case ModuleReadOnly:
                    current.IsReadOnly = true;
                    break;
                case ModulePrivate:
                    current.IsPrivate = true;
                    break;
                case ModuleTerminator:
                    current.Finish(modules);
                    break;
            }
        }

        current.Finish(modules);
        return new Contents(codePage, projectName, description, constants, references, modules);
    }

    /// <summary>
    /// Reads the library identifier out of a reference record. All three shapes open with a
    /// length-prefixed identifier, so only the kind and the name that preceded it differ.
    /// </summary>
    private static VbaReference ReadReference(
        ushort id, ReadOnlySpan<byte> body, Encoding text, ref string? pendingName)
    {
        string? name = pendingName;
        pendingName = null;

        // The first four bytes size the identifier that follows; the rest of the record is
        // reserved fields nothing reads.
        string libid = body.Length >= 4
            ? ReadSized(body, text)
            : string.Empty;

        VbaReferenceKind kind = id switch
        {
            ReferenceProject => VbaReferenceKind.Project,
            ReferenceControl => VbaReferenceKind.Control,
            _ => VbaReferenceKind.Library,
        };

        return VbaReference.Parse(name, kind, libid);
    }

    /// <summary>Reads a four-byte length followed by that many bytes of text.</summary>
    private static string ReadSized(ReadOnlySpan<byte> body, Encoding text)
    {
        int size = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        return size < 0 || size > body.Length - 4 ? string.Empty : text.GetString(body.Slice(4, size));
    }

    private static string? Trimmed(string value) => value.Length == 0 ? null : value;

    /// <summary>The encoding of the project's byte strings, or the fallback when it is unavailable.</summary>
    public static Encoding Resolve(int codePage, Encoding fallback)
    {
        if (codePage is 0 or 1200)
            return fallback;

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException)
        {
            // Most single-byte code pages need CodePagesEncodingProvider registered, which is
            // the caller's decision to make.
            return fallback;
        }
    }

    /// <summary>The module being described, filled in as its records go by.</summary>
    private sealed class ModuleBuilder
    {
        public string? Name { get; set; }

        public string? StreamName { get; set; }

        public int SourceOffset { get; set; }

        public bool IsProcedural { get; set; }

        public bool IsReadOnly { get; set; }

        public bool IsPrivate { get; set; }

        public string? Description { get; set; }

        public void Finish(List<VbaModuleEntry> into)
        {
            if (Name is { Length: > 0 } name)
            {
                into.Add(new VbaModuleEntry(
                    name, StreamName is { Length: > 0 } stream ? stream : name,
                    SourceOffset, IsProcedural, IsReadOnly, IsPrivate, Description));
            }

            Name = null;
            StreamName = null;
            SourceOffset = 0;
            IsProcedural = false;
            IsReadOnly = false;
            IsPrivate = false;
            Description = null;
        }
    }
}
