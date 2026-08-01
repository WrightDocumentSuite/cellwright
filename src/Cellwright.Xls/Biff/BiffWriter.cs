using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Biff;

/// <summary>
/// Builds a BIFF record stream. Records are written into a growing buffer; anything longer than
/// a record may be gets split across <c>Continue</c> records, which is the mirror of what
/// <see cref="BiffReader"/> undoes.
/// </summary>
internal sealed class BiffWriter : IDisposable
{
    private byte[] _buffer;
    private int _length;

    /// <summary>Creates a writer with room for a typical workbook to start with.</summary>
    public BiffWriter(int initialCapacity = 1 << 16)
    {
        _buffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
    }

    /// <summary>Bytes written so far.</summary>
    public int Length => _length;

    /// <summary>The stream as written.</summary>
    public ReadOnlyMemory<byte> Written => _buffer.AsMemory(0, _length);

    /// <summary>Writes a whole record, splitting it across continuations when it is too long.</summary>
    public void WriteRecord(ushort type, ReadOnlySpan<byte> body)
    {
        if (body.Length <= BiffReader.MaxRecordSize)
        {
            WriteHeader(type, body.Length);
            Append(body);
            return;
        }

        WriteHeader(type, BiffReader.MaxRecordSize);
        Append(body[..BiffReader.MaxRecordSize]);

        int written = BiffReader.MaxRecordSize;
        while (written < body.Length)
        {
            int take = Math.Min(BiffReader.MaxRecordSize, body.Length - written);
            WriteHeader(BiffRecordType.Continue, take);
            Append(body.Slice(written, take));
            written += take;
        }
    }

    /// <summary>Writes a record with no body, such as EOF.</summary>
    public void WriteRecord(ushort type) => WriteHeader(type, 0);

    /// <summary>
    /// Starts a record whose length is not known yet, returning the offset to close it at.
    /// Used where the body is built by appending, such as a run of merged ranges.
    /// </summary>
    public int BeginRecord(ushort type)
    {
        WriteHeader(type, 0);
        return _length;
    }

    /// <summary>Fills in the length of a record started with <see cref="BeginRecord"/>.</summary>
    public void EndRecord(int bodyStart)
    {
        int size = _length - bodyStart;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(bodyStart - 2), (ushort)size);
    }

    /// <summary>Appends raw bytes to the record being built.</summary>
    public void Append(ReadOnlySpan<byte> bytes)
    {
        EnsureCapacity(bytes.Length);
        bytes.CopyTo(_buffer.AsSpan(_length));
        _length += bytes.Length;
    }

    /// <summary>Appends one byte.</summary>
    public void AppendByte(byte value)
    {
        EnsureCapacity(1);
        _buffer[_length++] = value;
    }

    /// <summary>Appends a little-endian 16-bit value.</summary>
    public void AppendUInt16(int value)
    {
        EnsureCapacity(2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_length), (ushort)value);
        _length += 2;
    }

    /// <summary>Appends a little-endian 32-bit value.</summary>
    public void AppendUInt32(uint value)
    {
        EnsureCapacity(4);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_length), value);
        _length += 4;
    }

    /// <summary>Appends a little-endian double.</summary>
    public void AppendDouble(double value)
    {
        EnsureCapacity(8);
        BinaryPrimitives.WriteDoubleLittleEndian(_buffer.AsSpan(_length), value);
        _length += 8;
    }

    /// <summary>Appends the given number of zero bytes.</summary>
    public void AppendZeros(int count)
    {
        EnsureCapacity(count);
        _buffer.AsSpan(_length, count).Clear();
        _length += count;
    }

    /// <summary>Overwrites a 32-bit value already written, for a field fixed up later.</summary>
    public void PatchUInt32(int offset, uint value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(offset), value);

    /// <summary>Appends the current position as a placeholder and returns where it was written.</summary>
    public int ReserveUInt32()
    {
        int offset = _length;
        AppendUInt32(0);
        return offset;
    }

    /// <summary>
    /// Writes an <c>XLUnicodeString</c>: a 16-bit character count, a flags byte, then the
    /// characters, compressed to one byte each when they all fit in a byte.
    /// </summary>
    public void AppendUnicodeString(string text)
    {
        AppendUInt16(text.Length);
        AppendCharacters(text);
    }

    /// <summary>Writes a <c>ShortXLUnicodeString</c>, whose count is one byte.</summary>
    public void AppendShortUnicodeString(string text)
    {
        string clipped = text.Length > 255 ? text[..255] : text;
        AppendByte((byte)clipped.Length);
        AppendCharacters(clipped);
    }

    private void AppendCharacters(string text)
    {
        bool compressible = IsCompressible(text);
        AppendByte(compressible ? (byte)0x00 : (byte)0x01);
        if (compressible)
        {
            EnsureCapacity(text.Length);
            for (int i = 0; i < text.Length; i++)
                _buffer[_length + i] = (byte)text[i];
            _length += text.Length;
            return;
        }

        EnsureCapacity(text.Length * 2);
        int written = Encoding.Unicode.GetBytes(text, _buffer.AsSpan(_length));
        _length += written;
    }

    /// <summary>Whether every character fits in the one-byte form, which halves the string's cost.</summary>
    public static bool IsCompressible(string text)
    {
        foreach (char c in text)
        {
            if (c > 0xFF)
                return false;
        }

        return true;
    }

    private void WriteHeader(ushort type, int size)
    {
        EnsureCapacity(4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_length), type);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_length + 2), (ushort)size);
        _length += 4;
    }

    private void EnsureCapacity(int additional)
    {
        if (_length + additional <= _buffer.Length)
            return;

        int capacity = Math.Max(_buffer.Length * 2, _length + additional);
        byte[] grown = ArrayPool<byte>.Shared.Rent(capacity);
        _buffer.AsSpan(0, _length).CopyTo(grown);
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = grown;
    }

    /// <summary>Copies the stream out, which is what goes into the container.</summary>
    public byte[] ToArray() => _buffer.AsSpan(0, _length).ToArray();

    /// <inheritdoc/>
    public void Dispose()
    {
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = [];
        _length = 0;
    }
}
