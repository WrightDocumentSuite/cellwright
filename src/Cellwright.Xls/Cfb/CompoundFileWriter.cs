using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Cfb;

/// <summary>
/// Writes a compound file holding the named streams a legacy workbook needs: the BIFF stream
/// itself and, beside it, the property sets that carry the document's title and author.
/// </summary>
/// <remarks>
/// Sectors are 512 bytes throughout. A stream below the 4096-byte cutoff has to go in the mini
/// stream rather than in sectors of its own: readers honour the cutoff when deciding where to
/// look, so a small stream written the other way is one they cannot find. That is not an edge
/// case here — the property streams are always small and the workbook never is, so every file
/// this writes uses both mechanisms at once.
/// </remarks>
internal static class CompoundFileWriter
{
    private const int SectorSize = 512;
    private const int MiniSectorSize = 64;
    private const int MiniStreamCutoff = 4096;
    private const int DirectoryEntrySize = 128;
    private const int EntriesPerDirectorySector = SectorSize / DirectoryEntrySize;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint FatSectorMark = 0xFFFFFFFD;
    private const uint NoStream = 0xFFFFFFFF;

    /// <summary>One stream to put in the container.</summary>
    /// <param name="Name">Name of the stream in the root storage.</param>
    /// <param name="Content">Its bytes.</param>
    public readonly record struct NamedStream(string Name, ReadOnlyMemory<byte> Content);

    /// <summary>Wraps one stream's bytes in a container and writes it out.</summary>
    public static ValueTask WriteAsync(
        Stream destination, string streamName, ReadOnlyMemory<byte> content, CancellationToken cancellationToken) =>
        WriteAsync(destination, [new NamedStream(streamName, content)], cancellationToken);

    /// <summary>Writes a container holding every stream given, in the root storage.</summary>
    public static async ValueTask WriteAsync(
        Stream destination, IReadOnlyList<NamedStream> streams, CancellationToken cancellationToken)
    {
        Layout layout = Plan(streams);

        await destination.WriteAsync(BuildHeader(layout), cancellationToken).ConfigureAwait(false);

        foreach (NamedStream stream in streams)
        {
            if (stream.Content.Length >= MiniStreamCutoff)
            {
                await WritePaddedAsync(
                    destination, stream.Content, RoundUpToMultiple(stream.Content.Length, SectorSize), cancellationToken)
                    .ConfigureAwait(false);
            }
        }

        await destination.WriteAsync(BuildMiniStream(streams, layout), cancellationToken).ConfigureAwait(false);
        await destination.WriteAsync(Serialize(layout.MiniFat), cancellationToken).ConfigureAwait(false);
        await destination.WriteAsync(BuildDirectory(streams, layout), cancellationToken).ConfigureAwait(false);
        await destination.WriteAsync(Serialize(layout.Fat), cancellationToken).ConfigureAwait(false);
        if (layout.DifatSectors > 0)
            await destination.WriteAsync(BuildDifat(layout), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Where everything sits, once the tables that describe it have been sized.</summary>
    private sealed record Layout
    {
        /// <summary>
        /// Where each stream begins: a sector for a large one, a mini sector for a small one,
        /// and the end-of-chain marker for an empty one — which is why these are not indexes.
        /// </summary>
        public required uint[] StreamStart { get; init; }

        public required uint MiniStreamStart { get; init; }
        public required int MiniStreamSectors { get; init; }
        public required int MiniStreamLength { get; init; }
        public required uint MiniFatStart { get; init; }
        public required int MiniFatSectors { get; init; }
        public required int DirectoryStart { get; init; }
        public required int DirectorySectors { get; init; }
        public required int FirstFatSector { get; init; }
        public required int FatSectors { get; init; }
        public required int FirstDifatSector { get; init; }
        public required int DifatSectors { get; init; }
        public required uint[] Fat { get; init; }
        public required uint[] MiniFat { get; init; }
    }

    /// <summary>
    /// Works out the whole layout: where each stream starts, how large the allocation tables
    /// must be, and what the tables then say. Sizing the tables is a fixed point rather than a
    /// formula, because each table sector is itself something the FAT has to account for.
    /// </summary>
    private static Layout Plan(IReadOnlyList<NamedStream> streams)
    {
        var starts = new uint[streams.Count];
        int sector = 0;
        int miniOffset = 0;

        for (int i = 0; i < streams.Count; i++)
        {
            int length = streams[i].Content.Length;
            if (length == 0)
            {
                starts[i] = EndOfChain;
            }
            else if (length >= MiniStreamCutoff)
            {
                starts[i] = (uint)sector;
                sector += RoundUpToMultiple(length, SectorSize) / SectorSize;
            }
            else
            {
                // A small stream is placed by mini sector, not by sector.
                starts[i] = (uint)(miniOffset / MiniSectorSize);
                miniOffset += RoundUpToMultiple(length, MiniSectorSize);
            }
        }

        int largeSectors = sector;
        int miniStreamLength = miniOffset;
        int miniSectors = miniStreamLength / MiniSectorSize;
        int miniStreamSectors = RoundUpToMultiple(miniStreamLength, SectorSize) / SectorSize;
        int miniFatSectors = miniSectors == 0 ? 0 : RoundUpToMultiple(miniSectors, FatEntriesPerSector) / FatEntriesPerSector;
        int directorySectors = Math.Max(
            1, RoundUpToMultiple(streams.Count + 1, EntriesPerDirectorySector) / EntriesPerDirectorySector);

        int miniStreamStart = largeSectors;
        int miniFatStart = miniStreamStart + miniStreamSectors;
        int directoryStart = miniFatStart + miniFatSectors;
        int occupied = directoryStart + directorySectors;

        (int fatSectors, int difatSectors) = SizeAllocationTables(occupied);
        int firstFatSector = occupied;
        int firstDifatSector = firstFatSector + fatSectors;

        uint[] fat = new uint[fatSectors * FatEntriesPerSector];
        Array.Fill(fat, FreeSector);
        for (int i = 0; i < streams.Count; i++)
        {
            int length = streams[i].Content.Length;
            if (length >= MiniStreamCutoff)
                Chain(fat, (int)starts[i], RoundUpToMultiple(length, SectorSize) / SectorSize);
        }

        Chain(fat, miniStreamStart, miniStreamSectors);
        Chain(fat, miniFatStart, miniFatSectors);
        Chain(fat, directoryStart, directorySectors);
        for (int i = 0; i < fatSectors; i++)
            fat[firstFatSector + i] = FatSectorMark;
        for (int i = 0; i < difatSectors; i++)
            fat[firstDifatSector + i] = DifatSectorMark;

        uint[] miniFat = new uint[Math.Max(miniFatSectors, 0) * FatEntriesPerSector];
        Array.Fill(miniFat, FreeSector);
        for (int i = 0; i < streams.Count; i++)
        {
            int length = streams[i].Content.Length;
            if (length is > 0 and < MiniStreamCutoff)
                Chain(miniFat, (int)starts[i], RoundUpToMultiple(length, MiniSectorSize) / MiniSectorSize);
        }

        return new Layout
        {
            StreamStart = starts,
            MiniStreamStart = miniStreamSectors == 0 ? EndOfChain : (uint)miniStreamStart,
            MiniStreamSectors = miniStreamSectors,
            MiniStreamLength = miniStreamLength,
            MiniFatStart = miniFatSectors == 0 ? EndOfChain : (uint)miniFatStart,
            MiniFatSectors = miniFatSectors,
            DirectoryStart = directoryStart,
            DirectorySectors = directorySectors,
            FirstFatSector = firstFatSector,
            FatSectors = fatSectors,
            FirstDifatSector = firstDifatSector,
            DifatSectors = difatSectors,
            Fat = fat,
            MiniFat = miniFat,
        };
    }

    /// <summary>The mini stream: every small stream in turn, each padded to a whole mini sector.</summary>
    private static byte[] BuildMiniStream(IReadOnlyList<NamedStream> streams, Layout layout)
    {
        var bytes = new byte[layout.MiniStreamSectors * SectorSize];
        for (int i = 0; i < streams.Count; i++)
        {
            int length = streams[i].Content.Length;
            if (length is > 0 and < MiniStreamCutoff)
                streams[i].Content.Span.CopyTo(bytes.AsSpan((int)layout.StreamStart[i] * MiniSectorSize));
        }

        return bytes;
    }

    /// <summary>Entries a FAT sector holds.</summary>
    private const int FatEntriesPerSector = SectorSize / 4;

    /// <summary>FAT sector numbers the header itself carries.</summary>
    private const int DifatEntriesInHeader = 109;

    /// <summary>FAT sector numbers a DIFAT sector carries, the last slot being the next pointer.</summary>
    private const int DifatEntriesPerSector = FatEntriesPerSector - 1;

    private const uint DifatSectorMark = 0xFFFFFFFC;

    /// <summary>
    /// Works out how many FAT and DIFAT sectors a file needs. Each of them occupies a sector
    /// that the FAT then has to account for, so the answer is a fixed point rather than a
    /// formula; it settles in two or three passes.
    /// </summary>
    private static (int FatSectors, int DifatSectors) SizeAllocationTables(int dataSectors)
    {
        int fatSectors = 1;
        int difatSectors = 0;

        for (int pass = 0; pass < 8; pass++)
        {
            int total = dataSectors + fatSectors + difatSectors;
            int neededFat = Math.Max(1, (total + FatEntriesPerSector - 1) / FatEntriesPerSector);
            int neededDifat = neededFat <= DifatEntriesInHeader
                ? 0
                : ((neededFat - DifatEntriesInHeader) + DifatEntriesPerSector - 1) / DifatEntriesPerSector;

            if (neededFat == fatSectors && neededDifat == difatSectors)
                break;

            fatSectors = neededFat;
            difatSectors = neededDifat;
        }

        return (fatSectors, difatSectors);
    }

    /// <summary>Builds the DIFAT sectors listing the FAT sectors that did not fit in the header.</summary>
    private static byte[] BuildDifat(Layout layout)
    {
        var bytes = new byte[layout.DifatSectors * SectorSize];
        Span<byte> span = bytes;
        int nextFat = layout.FirstFatSector + DifatEntriesInHeader;
        int remaining = layout.FatSectors - DifatEntriesInHeader;

        for (int sector = 0; sector < layout.DifatSectors; sector++)
        {
            Span<byte> current = span.Slice(sector * SectorSize, SectorSize);
            for (int slot = 0; slot < DifatEntriesPerSector; slot++)
            {
                uint value = remaining > 0 ? (uint)nextFat : FreeSector;
                BinaryPrimitives.WriteUInt32LittleEndian(current[(slot * 4)..], value);
                if (remaining <= 0)
                    continue;

                nextFat++;
                remaining--;
            }

            uint next = sector + 1 < layout.DifatSectors ? (uint)(layout.FirstDifatSector + sector + 1) : EndOfChain;
            BinaryPrimitives.WriteUInt32LittleEndian(current[(DifatEntriesPerSector * 4)..], next);
        }

        return bytes;
    }

    /// <summary>Links a run of consecutive sectors into one chain ending in the end marker.</summary>
    private static void Chain(uint[] table, int first, int count)
    {
        for (int i = 0; i < count; i++)
            table[first + i] = i == count - 1 ? EndOfChain : (uint)(first + i + 1);
    }

    private static byte[] BuildHeader(Layout layout)
    {
        var header = new byte[SectorSize];
        Span<byte> span = header;
        CompoundFile.Signature.CopyTo(span);

        BinaryPrimitives.WriteUInt16LittleEndian(span[24..], 0x003E);      // Minor version.
        BinaryPrimitives.WriteUInt16LittleEndian(span[26..], 0x0003);      // Major version 3: 512-byte sectors.
        BinaryPrimitives.WriteUInt16LittleEndian(span[28..], 0xFFFE);      // Little-endian marker.
        BinaryPrimitives.WriteUInt16LittleEndian(span[30..], 9);           // Sector shift: 2^9 = 512.
        BinaryPrimitives.WriteUInt16LittleEndian(span[32..], 6);           // Mini sector shift: 2^6 = 64.
        BinaryPrimitives.WriteUInt32LittleEndian(span[44..], (uint)layout.FatSectors);
        BinaryPrimitives.WriteUInt32LittleEndian(span[48..], (uint)layout.DirectoryStart);
        BinaryPrimitives.WriteUInt32LittleEndian(span[56..], MiniStreamCutoff);
        BinaryPrimitives.WriteUInt32LittleEndian(span[60..], (uint)layout.MiniFatStart);
        BinaryPrimitives.WriteUInt32LittleEndian(span[64..], (uint)layout.MiniFatSectors);
        BinaryPrimitives.WriteUInt32LittleEndian(
            span[68..], layout.DifatSectors > 0 ? (uint)layout.FirstDifatSector : EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(span[72..], (uint)layout.DifatSectors);

        // The header itself lists the first 109 FAT sectors; any beyond that are in the DIFAT.
        for (int i = 0; i < DifatEntriesInHeader; i++)
        {
            uint value = i < layout.FatSectors ? (uint)(layout.FirstFatSector + i) : FreeSector;
            BinaryPrimitives.WriteUInt32LittleEndian(span[(76 + (i * 4))..], value);
        }

        return header;
    }

    /// <summary>
    /// Builds the directory: the root storage, which owns the mini stream, and one entry per
    /// stream hanging off it.
    /// </summary>
    /// <remarks>
    /// The entries form a red-black tree, and a reader is entitled to search it rather than
    /// walk it. Writing them in the order the format collates by — shorter names first, then
    /// case-insensitively — makes a chain of right children a valid search tree, which is the
    /// simplest shape that stays findable.
    /// </remarks>
    private static byte[] BuildDirectory(IReadOnlyList<NamedStream> streams, Layout layout)
    {
        int[] order = Enumerable.Range(0, streams.Count).ToArray();
        Array.Sort(order, (left, right) => Collate(streams[left].Name, streams[right].Name));

        var directory = new byte[layout.DirectorySectors * SectorSize];
        Span<byte> span = directory;

        WriteEntry(
            span[..DirectoryEntrySize],
            "Root Entry",
            kind: 5,
            startSector: (uint)layout.MiniStreamStart,
            length: (ulong)layout.MiniStreamLength,
            child: streams.Count > 0 ? 1u : NoStream,
            rightSibling: NoStream);

        for (int slot = 0; slot < order.Length; slot++)
        {
            int index = order[slot];
            Span<byte> entry = span.Slice((slot + 1) * DirectoryEntrySize, DirectoryEntrySize);
            WriteEntry(
                entry,
                streams[index].Name,
                kind: 2,
                startSector: (uint)layout.StreamStart[index],
                length: (ulong)streams[index].Content.Length,
                child: NoStream,
                rightSibling: slot + 1 < order.Length ? (uint)(slot + 2) : NoStream);
        }

        // Every slot the sectors have room for must be marked unused, or a reader walking the
        // tree will follow whatever the bytes happen to say.
        for (int slot = streams.Count + 1; slot < layout.DirectorySectors * EntriesPerDirectorySector; slot++)
        {
            Span<byte> entry = span.Slice(slot * DirectoryEntrySize, DirectoryEntrySize);
            entry[66] = 0; // Unused.
            BinaryPrimitives.WriteUInt32LittleEndian(entry[68..], NoStream);
            BinaryPrimitives.WriteUInt32LittleEndian(entry[72..], NoStream);
            BinaryPrimitives.WriteUInt32LittleEndian(entry[76..], NoStream);
        }

        return directory;
    }

    /// <summary>
    /// The order the format sorts directory entries in (MS-CFB 2.6.4): by the length of the
    /// name first, and only then by the name itself, compared without regard to case.
    /// </summary>
    private static int Collate(string left, string right)
    {
        int byLength = left.Length.CompareTo(right.Length);
        return byLength != 0 ? byLength : string.CompareOrdinal(left.ToUpperInvariant(), right.ToUpperInvariant());
    }

    private static void WriteEntry(
        Span<byte> entry, string name, byte kind, uint startSector, ulong length, uint child, uint rightSibling)
    {
        int written = Encoding.Unicode.GetBytes(name, entry[..62]);
        BinaryPrimitives.WriteUInt16LittleEndian(entry[64..], (ushort)(written + 2));
        entry[66] = kind;
        entry[67] = 1; // Black, which keeps the red-black tree trivially valid.
        BinaryPrimitives.WriteUInt32LittleEndian(entry[68..], NoStream);      // No left sibling.
        BinaryPrimitives.WriteUInt32LittleEndian(entry[72..], rightSibling);
        BinaryPrimitives.WriteUInt32LittleEndian(entry[76..], child);
        BinaryPrimitives.WriteUInt32LittleEndian(entry[116..], startSector);
        BinaryPrimitives.WriteUInt64LittleEndian(entry[120..], length);
    }

    private static byte[] Serialize(uint[] fat)
    {
        var bytes = new byte[fat.Length * 4];
        for (int i = 0; i < fat.Length; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(i * 4), fat[i]);
        return bytes;
    }

    private static async ValueTask WritePaddedAsync(
        Stream destination, ReadOnlyMemory<byte> content, int paddedLength, CancellationToken cancellationToken)
    {
        await destination.WriteAsync(content, cancellationToken).ConfigureAwait(false);
        int padding = paddedLength - content.Length;
        if (padding <= 0)
            return;

        byte[] zeros = new byte[padding];
        await destination.WriteAsync(zeros, cancellationToken).ConfigureAwait(false);
    }

    private static int RoundUpToMultiple(int value, int multiple) => (value + multiple - 1) / multiple * multiple;
}
