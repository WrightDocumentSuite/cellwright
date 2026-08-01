using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Builds a stream of BIFF12 records for a test to read back.
/// </summary>
/// <remarks>
/// This exists because the corpus does not cover everything the format can say. A fixture built
/// here proves only that the reader understands the bytes this file writes, so it is worth
/// exactly as much as the byte layouts being right — which is why each caller cites the section
/// of MS-XLSB it is spelling out. It is not a substitute for a file Excel wrote.
/// </remarks>
internal sealed class Biff12Builder
{
    private readonly List<byte> _bytes = [];

    public byte[] ToArray() => [.. _bytes];

    /// <summary>Appends one record: the type, the length, and the body (MS-XLSB 2.1.4).</summary>
    public Biff12Builder Record(int type, ReadOnlySpan<byte> body)
    {
        if (type < 0x80)
        {
            _bytes.Add((byte)type);
        }
        else
        {
            _bytes.Add((byte)((type & 0x7F) | 0x80));
            _bytes.Add((byte)((type >> 7) & 0x7F));
        }

        int size = body.Length;
        do
        {
            byte piece = (byte)(size & 0x7F);
            size >>= 7;
            _bytes.Add(size > 0 ? (byte)(piece | 0x80) : piece);
        }
        while (size > 0);

        _bytes.AddRange(body);
        return this;
    }

    /// <summary>Appends a record with no body, which is how a collection opens and closes.</summary>
    public Biff12Builder Record(int type) => Record(type, ReadOnlySpan<byte>.Empty);

    /// <summary>Collects the bytes of one record body.</summary>
    internal sealed class Body
    {
        private readonly List<byte> _bytes = [];

        public byte[] ToArray() => [.. _bytes];

        public Body UInt32(uint value)
        {
            Span<byte> buffer = stackalloc byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
            _bytes.AddRange(buffer);
            return this;
        }

        public Body Int32(int value) => UInt32(unchecked((uint)value));

        public Body UInt16(int value)
        {
            Span<byte> buffer = stackalloc byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(buffer, (ushort)value);
            _bytes.AddRange(buffer);
            return this;
        }

        public Body Byte(int value)
        {
            _bytes.Add((byte)value);
            return this;
        }

        /// <summary>A range as four zero-based indexes: first row, last row, first column, last.</summary>
        public Body Range(int firstRow, int lastRow, int firstColumn, int lastColumn) =>
            UInt32((uint)firstRow).UInt32((uint)lastRow).UInt32((uint)firstColumn).UInt32((uint)lastColumn);

        /// <summary>An <c>XLWideString</c>: a count of characters, then the characters.</summary>
        public Body Text(string value)
        {
            UInt32((uint)value.Length);
            _bytes.AddRange(Encoding.Unicode.GetBytes(value));
            return this;
        }

        /// <summary>An <c>XLNullableWideString</c> holding nothing.</summary>
        public Body NoText() => UInt32(0xFFFFFFFF);

        /// <summary>A <c>RichStr</c> with no runs: the flags byte, then the text.</summary>
        public Body RichText(string value) => Byte(0).Text(value);

        /// <summary>Pads out to a fixed size, for the reserved tails these records carry.</summary>
        public Body Pad(int count)
        {
            _bytes.AddRange(new byte[count]);
            return this;
        }
    }
}
