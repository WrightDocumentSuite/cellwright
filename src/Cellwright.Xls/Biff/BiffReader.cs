using System.Buffers;
using System.Buffers.Binary;

namespace Cellwright.Xls.Biff;

/// <summary>
/// Walks a BIFF stream record by record. A record is a four-byte header — type and size — and a
/// body of at most 8224 bytes; anything longer is split across <c>Continue</c> records, which
/// this reader glues back together so the parsers above never see the seam.
/// </summary>
/// <remarks>
/// The stream is a byte array the caller already holds, so a record body is normally a slice of
/// it with nothing copied. Only a continued record is assembled, into a pooled buffer that the
/// next record reuses.
/// </remarks>
internal sealed class BiffReader : IDisposable
{
    /// <summary>Largest body a single record may carry (MS-XLS 2.1.4).</summary>
    public const int MaxRecordSize = 8224;

    private readonly byte[] _stream;
    private readonly IRecordDecryptor? _decryptor;
    private byte[]? _assembled;
    private int _position;

    /// <summary>Creates a reader over the bytes of a workbook stream.</summary>
    /// <param name="stream">The whole <c>Workbook</c> or <c>Book</c> stream.</param>
    /// <param name="decryptor">Decrypts record bodies, or <see langword="null"/> for a plain file.</param>
    public BiffReader(byte[] stream, IRecordDecryptor? decryptor = null)
    {
        _stream = stream;
        _decryptor = decryptor;
    }

    /// <summary>Type of the record the reader is on.</summary>
    public ushort RecordType { get; private set; }

    /// <summary>Body of the record the reader is on, with any continuations already joined.</summary>
    public ReadOnlyMemory<byte> Record { get; private set; }

    /// <summary>Offset of the current record's header in the stream, which BoundSheet8 points at.</summary>
    public int RecordOffset { get; private set; }

    /// <summary>Offset just past the current record, where the next header starts.</summary>
    public int Position => _position;

    /// <summary>
    /// Whether the current record's continuations were joined. The SST needs to know: its
    /// strings carry a per-fragment encoding flag that only makes sense before joining.
    /// </summary>
    public bool WasContinued { get; private set; }

    /// <summary>Sizes of the pieces a continued record arrived in, in order.</summary>
    public IReadOnlyList<int> ContinuationSizes => _continuationSizes;

    private readonly List<int> _continuationSizes = [];

    /// <summary>Moves to the next record. Returns false at the end of the stream.</summary>
    public bool Read()
    {
        if (!TryReadHeader(_position, out ushort type, out int size))
            return false;

        RecordOffset = _position;
        RecordType = type;
        int bodyStart = _position + 4;
        _position = bodyStart + size;
        WasContinued = false;
        _continuationSizes.Clear();

        // A Continue that follows a record the reader does not join belongs to that record; it
        // is only reachable here when the previous record was skipped, and then it is noise.
        if (!JoinsContinuations(type))
        {
            Record = Decrypt(bodyStart, size, type);
            return true;
        }

        int totalSize = size;
        int scan = _position;
        while (TryReadHeader(scan, out ushort nextType, out int nextSize) && nextType == BiffRecordType.Continue)
        {
            totalSize += nextSize;
            scan += 4 + nextSize;
        }

        if (totalSize == size)
        {
            Record = Decrypt(bodyStart, size, type);
            return true;
        }

        WasContinued = true;
        Assemble(bodyStart, size, totalSize, type);
        return true;
    }

    /// <summary>Positions the reader so that the next <see cref="Read"/> starts at an offset.</summary>
    public void Seek(int offset) => _position = Math.Clamp(offset, 0, _stream.Length);

    /// <summary>Reads the type of the record at an offset without moving the reader.</summary>
    public ushort PeekTypeAt(int offset) => TryReadHeader(offset, out ushort type, out _) ? type : (ushort)0;

    private bool TryReadHeader(int offset, out ushort type, out int size)
    {
        type = 0;
        size = 0;
        if (offset < 0 || offset + 4 > _stream.Length)
            return false;

        type = BinaryPrimitives.ReadUInt16LittleEndian(_stream.AsSpan(offset));
        size = BinaryPrimitives.ReadUInt16LittleEndian(_stream.AsSpan(offset + 2));

        // A size that runs past the end means the stream was truncated; the record is refused
        // rather than read out of bounds.
        return offset + 4 + size <= _stream.Length;
    }

    /// <summary>
    /// Whether a record's continuations are joined into it. Most are, but a few carry
    /// independent payloads in their Continue records and are handled by their own parsers.
    /// </summary>
    private static bool JoinsContinuations(ushort type) => type != BiffRecordType.Continue;

    private void Assemble(int bodyStart, int firstSize, int totalSize, ushort type)
    {
        if (_assembled is null || _assembled.Length < totalSize)
        {
            if (_assembled is not null)
                ArrayPool<byte>.Shared.Return(_assembled);
            _assembled = ArrayPool<byte>.Shared.Rent(totalSize);
        }

        Decrypt(bodyStart, firstSize, type).Span.CopyTo(_assembled);
        _continuationSizes.Add(firstSize);

        int written = firstSize;
        int scan = _position;
        while (TryReadHeader(scan, out ushort nextType, out int nextSize) && nextType == BiffRecordType.Continue)
        {
            Decrypt(scan + 4, nextSize, nextType).Span.CopyTo(_assembled.AsSpan(written));
            _continuationSizes.Add(nextSize);
            written += nextSize;
            scan += 4 + nextSize;
        }

        _position = scan;
        Record = _assembled.AsMemory(0, written);
    }

    private ReadOnlyMemory<byte> Decrypt(int offset, int size, ushort type)
    {
        if (_decryptor is null)
            return _stream.AsMemory(offset, size);

        return _decryptor.Decrypt(_stream, offset, size, type);
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_assembled is null)
            return;

        ArrayPool<byte>.Shared.Return(_assembled);
        _assembled = null;
    }
}

/// <summary>
/// Turns the encrypted body of a record back into plain bytes. The cipher runs over the whole
/// stream, so a decryptor has to see every byte in order, including the ones it leaves alone.
/// </summary>
internal interface IRecordDecryptor
{
    /// <summary>Returns the plain bytes of one record body.</summary>
    /// <param name="stream">The whole workbook stream.</param>
    /// <param name="offset">Where the body starts.</param>
    /// <param name="size">How long the body is.</param>
    /// <param name="recordType">Type of the record, since a few are never encrypted.</param>
    public ReadOnlyMemory<byte> Decrypt(byte[] stream, int offset, int size, ushort recordType);
}
