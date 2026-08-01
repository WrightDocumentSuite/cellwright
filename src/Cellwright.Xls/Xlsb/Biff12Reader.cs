namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Walks the records of one part of a binary workbook.
/// </summary>
/// <remarks>
/// BIFF12 frames its records differently from the older BIFF8 (MS-XLSB 2.1.4): both the type
/// and the size are written seven bits at a time, so a common record costs two bytes of
/// overhead instead of four, and a record body is no longer capped at 8224 bytes. That last
/// part is the welcome one — there are no <c>Continue</c> records to reassemble, and a body is
/// always a single slice of the part with nothing copied.
/// </remarks>
internal sealed class Biff12Reader(byte[] part)
{
    private readonly byte[] _part = part;
    private int _position;
    private int _bodyStart;
    private int _bodySize;

    /// <summary>Type of the record the reader is on.</summary>
    public int RecordType { get; private set; } = -1;

    /// <summary>Body of the record the reader is on.</summary>
    public ReadOnlySpan<byte> Record => _part.AsSpan(_bodyStart, _bodySize);

    /// <summary>Moves to the next record. Returns false at the end of the part.</summary>
    public bool Read()
    {
        int cursor = _position;
        if (!TryReadType(ref cursor, out int type) || !TryReadSize(ref cursor, out int size))
            return false;

        // A size running past the end means the part was truncated; stopping is the whole of
        // the recovery available, since there is no way to find the next header without it.
        if (size < 0 || cursor + size > _part.Length)
            return false;

        RecordType = type;
        _bodyStart = cursor;
        _bodySize = size;
        _position = cursor + size;
        return true;
    }

    /// <summary>
    /// Reads on until the record that closes the collection the reader is inside, leaving the
    /// reader on it. Collections nest, so the matching end is the one at the same depth.
    /// </summary>
    /// <param name="begin">The record that opened the collection.</param>
    /// <param name="end">The record that closes it.</param>
    public void SkipCollection(int begin, int end)
    {
        int depth = 1;
        while (depth > 0 && Read())
        {
            if (RecordType == begin)
                depth++;
            else if (RecordType == end)
                depth--;
        }
    }

    /// <summary>
    /// The type is one byte, or two when the low byte has its high bit set; either way the
    /// value is the seven low bits of each, least significant first.
    /// </summary>
    private bool TryReadType(ref int cursor, out int type)
    {
        type = 0;
        if (cursor >= _part.Length)
            return false;

        int low = _part[cursor++];
        if ((low & 0x80) == 0)
        {
            type = low;
            return true;
        }

        if (cursor >= _part.Length)
            return false;

        type = (low & 0x7F) | ((_part[cursor++] & 0x7F) << 7);
        return true;
    }

    /// <summary>
    /// The size is one to four bytes of seven bits each, the high bit of every byte saying
    /// whether another follows. The fourth byte's flag is meaningless and ignored, which is
    /// what ends the loop after four rounds.
    /// </summary>
    private bool TryReadSize(ref int cursor, out int size)
    {
        size = 0;
        for (int shift = 0; shift < 28; shift += 7)
        {
            if (cursor >= _part.Length)
                return false;

            int piece = _part[cursor++];
            size |= (piece & 0x7F) << shift;
            if ((piece & 0x80) == 0)
                break;
        }

        return true;
    }
}
