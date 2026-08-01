using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Builds one part of a binary workbook, record by record.
/// </summary>
/// <remarks>
/// The mirror of <see cref="Biff12Reader"/>: a record is a type and a size, each written seven
/// bits at a time, then the body (MS-XLSB 2.1.4). Two things follow from that framing and shape
/// this class. A record's size has to be known before its header can be written, so a body is
/// built in a scratch buffer and copied in afterwards. And there is no 8224-byte cap and no
/// <c>Continue</c> record to work around, which is why nothing here has to split anything.
/// </remarks>
internal sealed class Biff12Writer : IDisposable
{
    private byte[] _buffer;
    private int _length;

    public Biff12Writer(int capacity = 4096) => _buffer = ArrayPool<byte>.Shared.Rent(capacity);

    /// <summary>How many bytes have been written.</summary>
    public int Length => _length;

    /// <summary>What has been written so far.</summary>
    public ReadOnlySpan<byte> Written => _buffer.AsSpan(0, _length);

    /// <summary>Writes a record with a body.</summary>
    /// <param name="type">The record type.</param>
    /// <param name="body">The record's body.</param>
    public void WriteRecord(int type, ReadOnlySpan<byte> body)
    {
        WriteType(type);
        WriteSize(body.Length);
        Append(body);
    }

    /// <summary>Writes a record that carries nothing, which is what the collection markers are.</summary>
    /// <param name="type">The record type.</param>
    public void WriteRecord(int type)
    {
        WriteType(type);
        WriteSize(0);
    }

    /// <summary>Appends raw bytes to the body being built.</summary>
    public void Append(scoped ReadOnlySpan<byte> bytes)
    {
        EnsureCapacity(bytes.Length);
        bytes.CopyTo(_buffer.AsSpan(_length));
        _length += bytes.Length;
    }

    public void AppendByte(byte value)
    {
        EnsureCapacity(1);
        _buffer[_length++] = value;
    }

    public void AppendUInt16(int value)
    {
        EnsureCapacity(2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_length), (ushort)value);
        _length += 2;
    }

    public void AppendUInt32(uint value)
    {
        EnsureCapacity(4);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_length), value);
        _length += 4;
    }

    public void AppendInt32(int value) => AppendUInt32((uint)value);

    public void AppendDouble(double value)
    {
        EnsureCapacity(8);
        BinaryPrimitives.WriteDoubleLittleEndian(_buffer.AsSpan(_length), value);
        _length += 8;
    }

    /// <summary>
    /// Writes an <c>XLWideString</c>: a count of UTF-16 code units, then the units themselves
    /// (MS-XLSB 2.5.166). Always two bytes per character — the compression the older format
    /// allowed for Latin text is gone.
    /// </summary>
    public void AppendString(string? text)
    {
        string value = text ?? string.Empty;
        AppendUInt32((uint)value.Length);
        AppendUtf16(value);
    }

    /// <summary>
    /// Writes an <c>XLNullableWideString</c>, whose count of <c>0xFFFFFFFF</c> means there is no
    /// string at all — which is different from a string of no characters.
    /// </summary>
    public void AppendNullableString(string? text)
    {
        if (text is null)
        {
            AppendUInt32(0xFFFFFFFF);
            return;
        }

        AppendString(text);
    }

    private void AppendUtf16(string value)
    {
        EnsureCapacity(value.Length * 2);
        int written = Encoding.Unicode.GetBytes(value, _buffer.AsSpan(_length));
        _length += written;
    }

    /// <summary>Empties the buffer without giving it back, so one writer can build many bodies.</summary>
    public void Reset() => _length = 0;

    /// <summary>
    /// The type is one byte when it fits in seven bits, and two otherwise, with the high bit of
    /// the first saying which.
    /// </summary>
    private void WriteType(int type)
    {
        if (type < 0x80)
        {
            AppendByte((byte)type);
            return;
        }

        AppendByte((byte)((type & 0x7F) | 0x80));
        AppendByte((byte)((type >> 7) & 0x7F));
    }

    /// <summary>
    /// The size is one to four bytes of seven bits each, the high bit of every byte but the last
    /// saying that another follows.
    /// </summary>
    private void WriteSize(int size)
    {
        uint remaining = (uint)size;
        for (int i = 0; i < 4; i++)
        {
            byte piece = (byte)(remaining & 0x7F);
            remaining >>= 7;

            if (remaining != 0 && i < 3)
                piece |= 0x80;

            AppendByte(piece);
            if (remaining == 0)
                return;
        }
    }

    private void EnsureCapacity(int needed)
    {
        if (_buffer.Length - _length >= needed)
            return;

        byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_buffer.Length * 2, _length + needed));
        _buffer.AsSpan(0, _length).CopyTo(bigger);
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = bigger;
    }

    /// <summary>The bytes written, as an array of their own.</summary>
    public byte[] ToArray() => _buffer.AsSpan(0, _length).ToArray();

    public void Dispose()
    {
        if (_buffer.Length == 0)
            return;

        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = [];
        _length = 0;
    }
}
