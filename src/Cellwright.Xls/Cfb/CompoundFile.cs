using System.Buffers.Binary;
using Cellwright;

namespace Cellwright.Xls.Cfb;

/// <summary>
/// A Compound File Binary container (MS-CFB) — the little FAT filesystem an .xls file is
/// wrapped in. Only what a workbook needs is implemented: the allocation tables, the directory,
/// and reading a named stream out.
/// </summary>
/// <remarks>
/// The whole file is held in memory. A .xls workbook cannot exceed the 65536×256 grid, so it is
/// bounded by construction, and the alternative — seeking a sector chain over a stream — buys
/// nothing when the BIFF layer above reads the stream end to end anyway.
/// </remarks>
internal sealed class CompoundFile
{
    /// <summary>The eight bytes that open every compound file.</summary>
    public static ReadOnlySpan<byte> Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private const int HeaderSize = 512;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint FatSector = 0xFFFFFFFD;
    private const uint DifatSector = 0xFFFFFFFC;
    private const int DirectoryEntrySize = 128;

    private readonly byte[] _data;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint[] _fat;
    private readonly uint[] _miniFat;
    private readonly List<DirectoryEntry> _directory;
    private readonly byte[] _miniStream;
    private readonly long _miniStreamCutoff;
    private readonly long _streamLimit;

    private CompoundFile(
        byte[] data, int sectorSize, int miniSectorSize, uint[] fat, uint[] miniFat,
        List<DirectoryEntry> directory, byte[] miniStream, long miniStreamCutoff, long streamLimit)
    {
        _data = data;
        _sectorSize = sectorSize;
        _miniSectorSize = miniSectorSize;
        _fat = fat;
        _miniFat = miniFat;
        _directory = directory;
        _miniStream = miniStream;
        _miniStreamCutoff = miniStreamCutoff;
        _streamLimit = streamLimit;
    }

    /// <summary>Whether the bytes start with the compound file signature.</summary>
    public static bool HasSignature(ReadOnlySpan<byte> head) =>
        head.Length >= Signature.Length && head[..Signature.Length].SequenceEqual(Signature);

    /// <summary>Parses a container that has already been read into memory.</summary>
    /// <exception cref="XlsFormatException">The header, allocation tables or directory are unusable.</exception>
    public static CompoundFile Parse(byte[] data, PackageLimits? limits = null)
    {
        limits ??= PackageLimits.Default;
        return Parse(data, limits.ApplyTo(DocumentLoadBudget.Default));
    }

    /// <summary>Parses a container under the shared spreadsheet load budget.</summary>
    internal static CompoundFile Parse(
        byte[] data, DocumentLoadBudget budget, DocumentLoadBudgetState? loadBudget = null)
    {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentNullException.ThrowIfNull(budget);
        loadBudget ??= new DocumentLoadBudgetState(budget);
        DocumentLoadBudgetState.Ensure(
            nameof(DocumentLoadBudget.MaxInputBytes), budget.MaxInputBytes, data.LongLength);

        if (!HasSignature(data))
            throw new XlsFormatException("The file does not start with the compound file signature.");
        if (data.Length < HeaderSize)
            throw new XlsFormatException("The compound file is shorter than its own header.");

        ReadOnlySpan<byte> header = data.AsSpan(0, HeaderSize);
        ushort majorVersion = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
        ushort byteOrder = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
        int sectorShift = BinaryPrimitives.ReadUInt16LittleEndian(header[30..]);
        int miniSectorShift = BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
        if (majorVersion is not (3 or 4))
            throw new XlsFormatException($"The compound file declares unsupported major version {majorVersion}.");
        if (byteOrder != 0xFFFE)
            throw new XlsFormatException("The compound file does not use the required little-endian byte order.");
        if ((majorVersion == 3 && sectorShift != 9) || (majorVersion == 4 && sectorShift != 12))
        {
            throw new XlsFormatException(
                $"Compound file version {majorVersion} requires sector shift " +
                $"{(majorVersion == 3 ? 9 : 12)}, not {sectorShift}.");
        }
        if (miniSectorShift != 6)
            throw new XlsFormatException($"The compound file declares mini-sector shift {miniSectorShift}; MS-CFB requires 6.");

        int sectorSize = 1 << sectorShift;
        int miniSectorSize = 1 << miniSectorShift;
        int availableSectors = data.Length <= sectorSize
            ? 0
            : (int)(((long)data.Length - sectorSize) / sectorSize);
        uint fatCountValue = BinaryPrimitives.ReadUInt32LittleEndian(header[44..]);
        uint firstDirectorySector = BinaryPrimitives.ReadUInt32LittleEndian(header[48..]);
        long miniStreamCutoff = BinaryPrimitives.ReadUInt32LittleEndian(header[56..]);
        uint firstMiniFatSector = BinaryPrimitives.ReadUInt32LittleEndian(header[60..]);
        uint miniFatCountValue = BinaryPrimitives.ReadUInt32LittleEndian(header[64..]);
        uint firstDifatSector = BinaryPrimitives.ReadUInt32LittleEndian(header[68..]);
        uint difatCountValue = BinaryPrimitives.ReadUInt32LittleEndian(header[72..]);

        int fatCount = BoundedSectorCount(fatCountValue, availableSectors, "FAT");
        int miniFatCount = BoundedSectorCount(miniFatCountValue, availableSectors, "mini FAT");
        int difatCount = BoundedSectorCount(difatCountValue, availableSectors, "DIFAT");

        // Every file Excel writes declares 4096, but the header is what decides, and a file
        // that says something else has to be believed or its small streams go missing.
        if (miniStreamCutoff is <= 0 or > int.MaxValue)
            miniStreamCutoff = DefaultMiniStreamCutoff;

        uint[] fat = ReadFat(
            data, sectorSize, header, firstDifatSector, difatCount, fatCount, availableSectors);
        List<DirectoryEntry> directory = ReadDirectory(
            data,
            sectorSize,
            fat,
            firstDirectorySector,
            majorVersion,
            budget.MaxPartBytes,
            budget.MaxPackageParts);
        uint[] miniFat = ReadChainAsSectorNumbers(
            data, sectorSize, fat, firstMiniFatSector, miniFatCount, budget.MaxPartBytes);

        // The mini stream lives inside the root entry, and every small stream is carved out of it.
        DirectoryEntry root = directory.Count > 0
            ? directory[0]
            : throw new XlsFormatException("The compound file has no root directory entry.");
        long miniStreamLength = CheckedStreamLength(root.Length, budget.MaxPartBytes, "root mini stream");
        byte[] miniStream = ReadChain(
            data, sectorSize, fat, root.StartSector, miniStreamLength, budget.MaxPartBytes, "root mini stream");

        var container = new CompoundFile(
            data, sectorSize, miniSectorSize, fat, miniFat, directory, miniStream, miniStreamCutoff, budget.MaxPartBytes);
        container.ValidateBudget(loadBudget);
        return container;
    }

    private void ValidateBudget(DocumentLoadBudgetState loadBudget)
    {
        loadBudget.AddPackageParts(_directory.Count);
        long inflated = 0;
        foreach (DirectoryEntry entry in _directory)
        {
            if (entry.Kind != DirectoryEntryKind.Stream)
                continue;

            long length = entry.Length > long.MaxValue ? long.MaxValue : (long)entry.Length;
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxPartBytes), loadBudget.Budget.MaxPartBytes, length);
            inflated = DocumentLoadBudgetState.AddWithoutOverflow(inflated, length);
        }
        loadBudget.AddInflatedBytes(inflated);

        int objectPool = FindIndex(["ObjectPool"]);
        if (objectPool < 0)
            return;

        foreach (int child in Subtree(_directory[objectPool].Child))
        {
            DirectoryEntry candidate = _directory[child];
            long bytes = candidate.Kind == DirectoryEntryKind.Stream
                ? candidate.Length > long.MaxValue ? long.MaxValue : (long)candidate.Length
                : DescendantStreamBytes(candidate.Child);
            loadBudget.AddEmbeddedObject(bytes);
        }
    }

    private long DescendantStreamBytes(uint child)
    {
        long total = 0;
        var pending = new Stack<uint>();
        var seen = new HashSet<int>();
        pending.Push(child);
        while (pending.TryPop(out uint siblingRoot))
        {
            foreach (int index in Subtree(siblingRoot))
            {
                if (!seen.Add(index))
                    continue;

                DirectoryEntry entry = _directory[index];
                if (entry.Kind == DirectoryEntryKind.Stream)
                {
                    total = DocumentLoadBudgetState.AddWithoutOverflow(
                        total, entry.Length > long.MaxValue ? long.MaxValue : (long)entry.Length);
                }
                else if (entry.Kind == DirectoryEntryKind.Storage)
                {
                    pending.Push(entry.Child);
                }
            }
        }

        return total;
    }

    /// <summary>Names of the streams and storages in the container, in directory order.</summary>
    public IEnumerable<string> EntryNames => _directory
        .Where(static entry => entry.Kind != DirectoryEntryKind.Empty)
        .Select(static entry => entry.Name);

    /// <summary>Reads a stream by name, or returns <see langword="null"/> when there is none.</summary>
    /// <remarks>
    /// The name is looked up across the whole directory, ignoring which storage holds it. That
    /// is what a workbook needs — its streams sit at the root and their names are distinctive —
    /// but a container whose storages reuse a name needs <see cref="TryReadPath"/> instead.
    /// </remarks>
    public byte[]? TryReadStream(string name)
    {
        foreach (DirectoryEntry entry in _directory)
        {
            if (entry.Kind == DirectoryEntryKind.Stream && string.Equals(entry.Name, name, StringComparison.Ordinal))
                return Read(entry);
        }

        return null;
    }

    /// <summary>
    /// Returns a named stream's declared length without materializing its sector chain.
    /// Format-specific readers use this for ceilings tighter than the general part budget.
    /// </summary>
    public long? TryGetStreamLength(string name)
    {
        foreach (DirectoryEntry entry in _directory)
        {
            if (entry.Kind == DirectoryEntryKind.Stream && string.Equals(entry.Name, name, StringComparison.Ordinal))
                return entry.Length > long.MaxValue ? long.MaxValue : (long)entry.Length;
        }

        return null;
    }

    /// <summary>
    /// Reads a stream by its path through the storages, e.g. <c>_VBA_PROJECT_CUR/VBA/dir</c>.
    /// </summary>
    /// <remarks>
    /// A VBA project names its module streams after the modules, so one can easily be called
    /// <c>Workbook</c> — the same name the BIFF stream of the very same container has. Only a
    /// walk that respects the storages tells the two apart.
    /// </remarks>
    public byte[]? TryReadPath(params ReadOnlySpan<string> path)
    {
        int found = FindIndex(path);
        return found >= 0 && _directory[found].Kind == DirectoryEntryKind.Stream ? Read(_directory[found]) : null;
    }

    /// <summary>Names of the entries directly inside a storage; the root when the path is empty.</summary>
    public IReadOnlyList<string> ChildNames(params ReadOnlySpan<string> storagePath)
    {
        int root = storagePath.IsEmpty ? 0 : FindIndex(storagePath);
        if (root < 0 || _directory[root].Kind is not (DirectoryEntryKind.Storage or DirectoryEntryKind.Root))
            return [];

        List<string> names = [];
        foreach (int child in Subtree(_directory[root].Child))
            names.Add(_directory[child].Name);
        return names;
    }

    /// <summary>Whether a storage of that path exists.</summary>
    public bool StorageExists(params ReadOnlySpan<string> path) =>
        FindIndex(path) is int found && found >= 0 && _directory[found].Kind == DirectoryEntryKind.Storage;

    /// <summary>
    /// Walks the directory from the root, one path segment at a time. Everything below the root
    /// is reached through it, so a name only has to be unique among its own siblings.
    /// </summary>
    private int FindIndex(scoped ReadOnlySpan<string> path)
    {
        if (_directory.Count == 0 || path.IsEmpty)
            return -1;

        uint children = _directory[0].Child;
        int found = -1;
        foreach (string segment in path)
        {
            found = FindSibling(children, segment);
            if (found < 0)
                return -1;
            children = _directory[found].Child;
        }

        return found;
    }

    /// <summary>Finds one name among a storage's children.</summary>
    private int FindSibling(uint child, string name)
    {
        foreach (int index in Subtree(child))
        {
            if (string.Equals(_directory[index].Name, name, StringComparison.OrdinalIgnoreCase))
                return index;
        }

        return -1;
    }

    /// <summary>
    /// Every entry of one sibling tree. The tree is ordered by name length and then by an
    /// uppercased comparison, but a file whose ordering is wrong is still a file somebody wants
    /// opened, so the whole subtree is walked rather than searched.
    /// </summary>
    private List<int> Subtree(uint root)
    {
        List<int> found = [];
        var pending = new Stack<uint>();
        var seen = new HashSet<uint>();
        pending.Push(root);
        while (pending.TryPop(out uint next))
        {
            if (next >= (uint)_directory.Count || !seen.Add(next))
                continue;

            DirectoryEntry entry = _directory[(int)next];
            if (entry.Kind != DirectoryEntryKind.Empty)
                found.Add((int)next);
            pending.Push(entry.Left);
            pending.Push(entry.Right);
        }

        return found;
    }

    private byte[] Read(DirectoryEntry entry)
    {
        int length = checked((int)CheckedStreamLength(entry.Length, _streamLimit, $"stream '{entry.Name}'"));

        // Anything under the cutoff is packed into the mini stream instead of taking a whole
        // sector of its own.
        return entry.Length < (ulong)_miniStreamCutoff
            ? ReadMiniChain(entry.StartSector, length)
            : ReadChain(_data, _sectorSize, _fat, entry.StartSector, length, _streamLimit, $"stream '{entry.Name}'");
    }

    /// <summary>What a header that says nothing sensible is taken to mean (MS-CFB 2.6.1).</summary>
    private const long DefaultMiniStreamCutoff = 4096;

    private static uint[] ReadFat(
        byte[] data,
        int sectorSize,
        ReadOnlySpan<byte> header,
        uint firstDifatSector,
        int difatCount,
        int fatSectorCount,
        int availableSectors)
    {
        // The first 109 FAT sector numbers sit in the header; the rest hang off a chain of DIFAT
        // sectors, each ending with a pointer to the next. Both lists are untrusted allocation
        // structures, so declared counts and repeated sectors are rejected before allocating.
        var fatSectors = new List<uint>(fatSectorCount);
        var seenFatSectors = new HashSet<uint>();
        for (int i = 0; i < 109 && fatSectors.Count < fatSectorCount; i++)
        {
            uint sector = BinaryPrimitives.ReadUInt32LittleEndian(header[(76 + (i * 4))..]);
            if (sector is FreeSector or EndOfChain)
                break;
            AddFatSector(sector);
        }

        uint difat = firstDifatSector;
        int perSector = (sectorSize / 4) - 1;
        var seenDifatSectors = new HashSet<uint>();
        int difatSectorsRead = 0;
        while (difatSectorsRead < difatCount)
        {
            if (difat is FreeSector or EndOfChain)
            {
                throw new XlsFormatException(
                    $"The compound file declares {difatCount} DIFAT sectors but its chain ends after {difatSectorsRead}.");
            }
            if (!seenDifatSectors.Add(difat))
                throw new XlsFormatException("The compound file contains a cyclic DIFAT chain.");
            if (difat >= (uint)availableSectors)
                throw new XlsFormatException($"The DIFAT refers to sector {difat}, which lies past the end of the file.");

            ReadOnlySpan<byte> sector = SectorSpan(data, sectorSize, difat);
            if (sector.Length < sectorSize)
                throw new XlsFormatException("The compound file ends in the middle of a DIFAT sector.");

            for (int i = 0; i < perSector && fatSectors.Count < fatSectorCount; i++)
            {
                uint entry = BinaryPrimitives.ReadUInt32LittleEndian(sector[(i * 4)..]);
                if (entry is FreeSector or EndOfChain)
                    continue;
                AddFatSector(entry);
            }

            difat = BinaryPrimitives.ReadUInt32LittleEndian(sector[(perSector * 4)..]);
            difatSectorsRead++;
        }

        if (difat is not (FreeSector or EndOfChain))
        {
            throw new XlsFormatException(
                $"The DIFAT chain continues past its declared {difatCount} sectors.");
        }

        if (fatSectors.Count != fatSectorCount)
        {
            throw new XlsFormatException(
                $"The compound file declares {fatSectorCount} FAT sectors but names only {fatSectors.Count}.");
        }

        var fat = new uint[checked(fatSectors.Count * (sectorSize / 4))];
        Array.Fill(fat, FreeSector);

        int position = 0;
        foreach (uint sectorNumber in fatSectors)
        {
            ReadOnlySpan<byte> sector = SectorSpan(data, sectorSize, sectorNumber);
            if (sector.Length < sectorSize)
                throw new XlsFormatException("The compound file ends in the middle of a FAT sector.");
            for (int i = 0; i < sectorSize / 4; i++)
                fat[position + i] = BinaryPrimitives.ReadUInt32LittleEndian(sector[(i * 4)..]);
            position += sectorSize / 4;
        }

        return fat;

        void AddFatSector(uint sector)
        {
            if (sector >= (uint)availableSectors)
                throw new XlsFormatException($"The FAT refers to sector {sector}, which lies past the end of the file.");
            if (!seenFatSectors.Add(sector))
                throw new XlsFormatException($"The compound file names FAT sector {sector} more than once.");
            fatSectors.Add(sector);
        }
    }

    private static List<DirectoryEntry> ReadDirectory(
        byte[] data,
        int sectorSize,
        uint[] fat,
        uint firstSector,
        ushort majorVersion,
        long streamLimit,
        int maxEntries)
    {
        long entryByteLimit = checked((long)maxEntries * DirectoryEntrySize);
        long effectiveLimit = Math.Min(streamLimit, entryByteLimit);
        byte[] bytes;
        try
        {
            bytes = ReadChain(data, sectorSize, fat, firstSector, length: -1, effectiveLimit, "directory");
        }
        catch (DocumentLoadLimitException) when (entryByteLimit < streamLimit)
        {
            throw new DocumentLoadLimitException(
                nameof(DocumentLoadBudget.MaxPackageParts), maxEntries, (long)maxEntries + 1);
        }

        DocumentLoadBudgetState.Ensure(
            nameof(DocumentLoadBudget.MaxPackageParts), maxEntries, bytes.Length / DirectoryEntrySize);
        var entries = new List<DirectoryEntry>(bytes.Length / DirectoryEntrySize);
        for (int offset = 0; offset + DirectoryEntrySize <= bytes.Length; offset += DirectoryEntrySize)
        {
            ReadOnlySpan<byte> entry = bytes.AsSpan(offset, DirectoryEntrySize);
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[64..]);
            var kind = (DirectoryEntryKind)entry[66];

            // Unused slots stay in the list: entries refer to their children and siblings by
            // position in the directory, and dropping one would renumber everything after it.

            // The length counts bytes including the terminating NUL of the UTF-16 name.
            nameLength = Math.Clamp(nameLength, 0, 64);
            string name = nameLength <= 2 ? "" : System.Text.Encoding.Unicode.GetString(entry[..(nameLength - 2)]);
            // Version 3 writers historically left the upper DWORD uninitialized. MS-CFB
            // requires current readers to ignore it rather than turn an ordinary stream huge.
            ulong length = majorVersion == 3
                ? BinaryPrimitives.ReadUInt32LittleEndian(entry[120..])
                : BinaryPrimitives.ReadUInt64LittleEndian(entry[120..]);
            entries.Add(new DirectoryEntry(
                name,
                kind,
                BinaryPrimitives.ReadUInt32LittleEndian(entry[116..]),
                length,
                BinaryPrimitives.ReadUInt32LittleEndian(entry[68..]),
                BinaryPrimitives.ReadUInt32LittleEndian(entry[72..]),
                BinaryPrimitives.ReadUInt32LittleEndian(entry[76..])));
        }

        return entries;
    }

    /// <summary>Reads a sector chain; a length of -1 means "to the end of the chain".</summary>
    private static byte[] ReadChain(
        byte[] data, int sectorSize, uint[] fat, uint start, long length, long sizeLimit, string description)
    {
        if (start is FreeSector or EndOfChain || length == 0)
            return [];
        if (length < -1)
            throw new ArgumentOutOfRangeException(nameof(length));
        if (length > sizeLimit || length > int.MaxValue)
            throw TooLarge(description, length, Math.Min(sizeLimit, int.MaxValue));

        var chunks = new List<uint>();
        var seen = new HashSet<uint>();
        uint sector = start;
        while (sector is not (FreeSector or EndOfChain or FatSector or DifatSector))
        {
            if (sector >= (uint)fat.Length)
                throw new XlsFormatException($"The {description} chain refers to sector {sector}, outside the FAT.");
            if (!seen.Add(sector))
                throw new XlsFormatException($"The compound file contains a cyclic {description} sector chain.");

            chunks.Add(sector);
            long available = (long)chunks.Count * sectorSize;
            if (available > sizeLimit || available > int.MaxValue)
                throw TooLarge(description, available, Math.Min(sizeLimit, int.MaxValue));
            if (length >= 0 && available >= length)
                break;

            sector = fat[sector];
        }

        long chainLength = (long)chunks.Count * sectorSize;
        if (length >= 0 && chainLength < length)
        {
            throw new XlsFormatException(
                $"The {description} declares {length} bytes but its sector chain holds only {chainLength}.");
        }

        int total = checked((int)(length >= 0 ? length : chainLength));
        var result = new byte[total];
        int written = 0;
        foreach (uint number in chunks)
        {
            ReadOnlySpan<byte> source = SectorSpan(data, sectorSize, number);
            int take = Math.Min(source.Length, total - written);
            if (take <= 0 || source.Length < Math.Min(sectorSize, total - written))
                throw new XlsFormatException($"The compound file ends in the middle of the {description} chain.");
            source[..take].CopyTo(result.AsSpan(written));
            written += take;
        }

        return result;
    }

    /// <summary>Reads a chain of sectors whose content is a table of little-endian sector numbers.</summary>
    private static uint[] ReadChainAsSectorNumbers(
        byte[] data, int sectorSize, uint[] fat, uint start, int sectorCount, long sizeLimit)
    {
        long length = checked((long)sectorCount * sectorSize);
        byte[] bytes = ReadChain(data, sectorSize, fat, start, length, sizeLimit, "mini FAT");
        var numbers = new uint[bytes.Length / 4];
        for (int i = 0; i < numbers.Length; i++)
            numbers[i] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4));
        return numbers;
    }

    private byte[] ReadMiniChain(uint start, int length)
    {
        if (length == 0)
            return [];

        var sectors = new List<uint>();
        var seen = new HashSet<uint>();
        uint sector = start;
        while (sector is not (FreeSector or EndOfChain) && (long)sectors.Count * _miniSectorSize < length)
        {
            if (sector >= (uint)_miniFat.Length)
                throw new XlsFormatException($"A mini stream chain refers to sector {sector}, outside the mini FAT.");
            if (!seen.Add(sector))
                throw new XlsFormatException("The compound file contains a cyclic mini stream sector chain.");
            sectors.Add(sector);
            sector = _miniFat[sector];
        }

        long available = (long)sectors.Count * _miniSectorSize;
        if (available < length)
            throw new XlsFormatException($"A mini stream declares {length} bytes but its sector chain holds only {available}.");

        var result = new byte[length];
        int written = 0;
        foreach (uint number in sectors)
        {
            int offset = checked((int)number * _miniSectorSize);
            if (offset < 0 || offset >= _miniStream.Length)
                throw new XlsFormatException($"Mini stream sector {number} lies outside the root mini stream.");
            int take = Math.Min(Math.Min(_miniSectorSize, length - written), _miniStream.Length - offset);
            if (take <= 0)
                throw new XlsFormatException("The root mini stream is shorter than a child stream declares.");
            _miniStream.AsSpan(offset, take).CopyTo(result.AsSpan(written));
            written += take;
        }

        return result;
    }

    /// <summary>The bytes of one sector; sector 0 starts after the header sector.</summary>
    private static ReadOnlySpan<byte> SectorSpan(byte[] data, int sectorSize, uint sector)
    {
        long offset = ((long)sector + 1) * sectorSize;
        if (offset < 0 || offset >= data.Length)
            return [];

        int available = (int)Math.Min(sectorSize, data.Length - offset);
        return data.AsSpan((int)offset, available);
    }

    private static int BoundedSectorCount(uint value, int availableSectors, string table)
    {
        if (value > (uint)availableSectors)
        {
            throw new XlsFormatException(
                $"The compound file declares {value} {table} sectors, but only {availableSectors} sectors fit in the file.");
        }

        return (int)value;
    }

    private static long CheckedStreamLength(ulong length, long limit, string description)
    {
        long effectiveLimit = Math.Min(limit, int.MaxValue);
        if (length > (ulong)Math.Max(0, effectiveLimit))
        {
            long observed = length > long.MaxValue ? long.MaxValue : (long)length;
            throw new DocumentLoadLimitException(
                nameof(DocumentLoadBudget.MaxPartBytes), effectiveLimit, observed);
        }
        return (long)length;
    }

    private static DocumentLoadLimitException TooLarge(string _, long size, long limit) => new(
        nameof(DocumentLoadBudget.MaxPartBytes), limit, size);

    private enum DirectoryEntryKind : byte
    {
        Empty = 0,
        Storage = 1,
        Stream = 2,
        Root = 5,
    }

    /// <param name="Name">Name of the stream or storage.</param>
    /// <param name="Kind">What the entry is.</param>
    /// <param name="StartSector">First sector of the content.</param>
    /// <param name="Length">Length of the content in bytes.</param>
    /// <param name="Left">Directory index of the left sibling, or <c>0xFFFFFFFF</c>.</param>
    /// <param name="Right">Directory index of the right sibling, or <c>0xFFFFFFFF</c>.</param>
    /// <param name="Child">Directory index of the first child of a storage, or <c>0xFFFFFFFF</c>.</param>
    private readonly record struct DirectoryEntry(
        string Name, DirectoryEntryKind Kind, uint StartSector, ulong Length, uint Left, uint Right, uint Child);
}
