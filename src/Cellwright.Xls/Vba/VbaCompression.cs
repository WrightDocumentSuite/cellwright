using System.Buffers.Binary;

namespace Cellwright.Xls.Vba;

/// <summary>
/// The compression a VBA project stores its source and its directory under (MS-OVBA 2.4.1).
/// A container is a run of chunks, each holding at most 4096 decompressed bytes; inside a chunk
/// a flag byte says which of the next eight tokens are literal bytes and which are back
/// references into what has already been decompressed.
/// </summary>
/// <remarks>
/// The awkward part, and where implementations of this usually go wrong, is that a copy token
/// splits its sixteen bits between length and offset according to how much of the
/// <em>current chunk</em> has been decompressed so far. The split therefore changes as the
/// chunk fills, and it resets at every chunk boundary.
/// </remarks>
internal static class VbaCompression
{
    /// <summary>The byte every compressed container opens with.</summary>
    private const byte ContainerSignature = 0x01;

    /// <summary>Decompressed bytes one chunk can hold.</summary>
    private const int ChunkSize = 4096;

    /// <summary>Bits 12 to 14 of a chunk header, which the specification fixes at 0b011.</summary>
    private const int HeaderSignature = 0x3000;

    private const int HeaderSignatureMask = 0x7000;
    private const int HeaderCompressedFlag = 0x8000;
    private const int HeaderSizeMask = 0x0FFF;

    /// <summary>
    /// How far a stream is allowed to expand. Six bytes of chunk can produce four thousand, so
    /// a file built for the purpose turns megabytes into gigabytes; a real project's source is
    /// a few hundred kilobytes, and this is two orders of magnitude above that.
    /// </summary>
    public const int DefaultSizeLimit = 64 * 1024 * 1024;

    /// <summary>Unpacks a compressed container.</summary>
    /// <param name="container">The container, starting at its signature byte.</param>
    /// <param name="sizeLimit">How many bytes the result may reach before the read is abandoned.</param>
    /// <exception cref="XlsFormatException">
    /// The container is not one, is truncated mid-chunk, or expands past <paramref name="sizeLimit"/>.
    /// </exception>
    public static byte[] Decompress(ReadOnlySpan<byte> container, int sizeLimit = DefaultSizeLimit) =>
        Decompress(container, loadBudget: null, CancellationToken.None, sizeLimit);

    /// <summary>Unpacks while contributing expansion to an outer document load.</summary>
    /// <param name="container">The container, starting at its signature byte.</param>
    /// <param name="loadBudget">Shared counters inherited from the workbook or package.</param>
    /// <param name="cancellationToken">Cancels decompression between tokens.</param>
    /// <param name="sizeLimit">The codec-specific ceiling in addition to the document budget.</param>
    public static byte[] Decompress(
        ReadOnlySpan<byte> container,
        DocumentLoadBudgetState? loadBudget,
        CancellationToken cancellationToken,
        int sizeLimit = DefaultSizeLimit)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (container.Length < 1 || container[0] != ContainerSignature)
            throw new XlsFormatException("The VBA stream does not begin with a compressed container.");

        // Starting small matters for a hostile compressed stream: its declared CFB length has
        // already been checked, but it must not turn into an equally large eager output buffer.
        var output = new List<byte>(Math.Min(container.Length, 64 * 1024));
        int position = 1;
        while (position + 1 < container.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int header = BinaryPrimitives.ReadUInt16LittleEndian(container[position..]);
            position += 2;

            if ((header & HeaderSignatureMask) != HeaderSignature)
                throw new XlsFormatException("A chunk of the VBA stream carries the wrong signature.");

            // The recorded size is three less than the whole chunk, header included, so the
            // data after the header runs one byte longer than the number itself.
            int end = Math.Min(position + (header & HeaderSizeMask) + 1, container.Length);
            int beforeChunk = output.Count;
            if ((header & HeaderCompressedFlag) == 0)
            {
                // A chunk that would not have compressed is stored as it is.
                foreach (byte value in container[position..Math.Min(position + ChunkSize, end)])
                {
                    EnsureCanGrow(output.Count, 1, sizeLimit, loadBudget);
                    output.Add(value);
                }
            }
            else
            {
                DecompressChunk(
                    container[position..end], output, sizeLimit, loadBudget, cancellationToken);
            }

            loadBudget?.AddInflatedBytes(output.Count - beforeChunk);

            position = end;
        }

        return [.. output];
    }

    private static XlsFormatException TooLarge(int sizeLimit) => new(
        $"The VBA stream expands past {sizeLimit} bytes, which no real project reaches. It was not read further.");

    private static void DecompressChunk(
        ReadOnlySpan<byte> chunk,
        List<byte> output,
        int sizeLimit,
        DocumentLoadBudgetState? loadBudget,
        CancellationToken cancellationToken)
    {
        int chunkStart = output.Count;
        int position = 0;
        while (position < chunk.Length)
        {
            byte flags = chunk[position++];
            for (int token = 0; token < 8 && position < chunk.Length; token++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if ((flags & (1 << token)) == 0)
                {
                    EnsureCanGrow(output.Count, 1, sizeLimit, loadBudget);
                    output.Add(chunk[position++]);
                    continue;
                }

                if (position + 1 >= chunk.Length)
                    return;

                Copy(
                    BinaryPrimitives.ReadUInt16LittleEndian(chunk[position..]),
                    output,
                    chunkStart,
                    sizeLimit,
                    loadBudget);
                position += 2;
            }
        }
    }

    /// <summary>
    /// Replays a back reference. The copy runs one byte at a time on purpose: a reference may
    /// reach into what it is itself producing, which is how a run of repeated bytes is encoded.
    /// </summary>
    /// <remarks>
    /// A reference is confined to its own chunk (MS-OVBA 2.4.1.3.19): it may not reach back past
    /// <c>DecompressedChunkStart</c>. One that does means the stream is not what it says it is,
    /// and carrying on would produce not text but a convincing imitation of it.
    /// </remarks>
    private static void Copy(
        int token,
        List<byte> output,
        int chunkStart,
        int sizeLimit,
        DocumentLoadBudgetState? loadBudget)
    {
        int written = output.Count - chunkStart;
        int offsetBits = Math.Max(4, CeilingLog2(written));
        int lengthMask = 0xFFFF >> offsetBits;

        int length = (token & lengthMask) + 3;
        int offset = ((token & ~lengthMask) >> (16 - offsetBits)) + 1;
        if (length >= ChunkSize || offset >= ChunkSize)
            throw new XlsFormatException($"A back reference of {length} bytes at {offset} back does not fit a chunk.");

        int source = output.Count - offset;
        if (source < chunkStart)
            throw new XlsFormatException("A back reference in the VBA stream points before the start of its chunk.");

        EnsureCanGrow(output.Count, length, sizeLimit, loadBudget);
        for (int i = 0; i < length; i++)
            output.Add(output[source + i]);
    }

    private static void EnsureCanGrow(
        int current,
        int additional,
        int sizeLimit,
        DocumentLoadBudgetState? loadBudget)
    {
        long observed = (long)current + additional;
        if (loadBudget is not null)
        {
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxPartBytes),
                loadBudget.Budget.MaxPartBytes,
                observed);
        }

        if (observed > sizeLimit)
            throw TooLarge(sizeLimit);
    }

    /// <summary>The smallest power of two that reaches <paramref name="value"/>, as an exponent.</summary>
    private static int CeilingLog2(int value)
    {
        int bits = 0;
        while (1 << bits < value)
            bits++;
        return bits;
    }
}
