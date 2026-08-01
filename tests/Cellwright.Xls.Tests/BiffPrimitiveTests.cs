using System.Text;
using Cellwright.Xls.Biff;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The encodings BIFF uses below the record level: the four-byte number, the strings, and the
/// splitting of a record too long to fit.
/// </summary>
public sealed class BiffPrimitiveTests
{
    [Theory]
    [InlineData(0x00000000u, 0.0)]
    [InlineData(0x3FF00000u, 1.0)]           // A double's top 30 bits.
    [InlineData(0x3FF00001u, 0.01)]          // The same, with the "divide by 100" bit set.
    [InlineData(0x00000002u, 0.0)]           // The integer form.
    [InlineData(0x00000006u, 1.0)]           // Integer 1, shifted up past the two flag bits.
    public void RkDecodesTheFourFormats(uint encoded, double expected)
    {
        double actual = BiffValues.DecodeRk(encoded);
        Assert.Equal(expected, actual, precision: 10);
    }

    [Fact]
    public void RkDecodesASmallInteger()
    {
        // 1234 shifted left two bits with the integer flag set.
        uint encoded = (1234u << 2) | 0x02;
        Assert.Equal(1234.0, BiffValues.DecodeRk(encoded));
    }

    [Fact]
    public void RkDecodesHundredths()
    {
        // 1234 hundredths with both the integer and the divide flags set.
        uint encoded = (1234u << 2) | 0x03;
        Assert.Equal(12.34, BiffValues.DecodeRk(encoded), precision: 10);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(1.0)]
    [InlineData(-1.0)]
    [InlineData(1234.0)]
    [InlineData(12.34)]
    [InlineData(0.25)]
    [InlineData(-536870912.0)]   // The lowest the 30-bit integer form holds.
    [InlineData(536870911.0)]
    public void RkRoundTripsEveryValueItAccepts(double value)
    {
        Assert.True(BiffValues.TryEncodeRk(value, out uint encoded), $"{value} should encode as RK.");
        Assert.Equal(value, BiffValues.DecodeRk(encoded));
    }

    [Theory]
    [InlineData(1.0 / 3.0)]
    [InlineData(1e300)]
    [InlineData(0.123456789)]
    public void RkRefusesWhatItCannotHoldExactly(double value)
    {
        if (BiffValues.TryEncodeRk(value, out uint encoded))
            Assert.Equal(value, BiffValues.DecodeRk(encoded));
    }

    [Fact]
    public void CompressedStringsReadBackAsWritten()
    {
        using var writer = new BiffWriter(64);
        writer.AppendUnicodeString("Hello");

        int offset = 0;
        Assert.Equal("Hello", BiffStrings.ReadUnicodeString(writer.Written.Span, ref offset));
        Assert.Equal(writer.Length, offset);
    }

    [Fact]
    public void WideStringsReadBackAsWritten()
    {
        using var writer = new BiffWriter(64);
        writer.AppendUnicodeString("Привет");

        int offset = 0;
        Assert.Equal("Привет", BiffStrings.ReadUnicodeString(writer.Written.Span, ref offset));
    }

    [Fact]
    public void ShortStringsReadBackAsWritten()
    {
        using var writer = new BiffWriter(64);
        writer.AppendShortUnicodeString("Sheet1");

        int offset = 0;
        Assert.Equal("Sheet1", BiffStrings.ReadShortUnicodeString(writer.Written.Span, ref offset));
    }

    [Fact]
    public void ByteStringsUseTheDeclaredCodePage()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1251 = Encoding.GetEncoding(1251);
        byte[] text = windows1251.GetBytes("Привет");

        var body = new byte[text.Length + 1];
        body[0] = (byte)text.Length;
        text.CopyTo(body, 1);

        int offset = 0;
        Assert.Equal("Привет", BiffStrings.ReadByteString(body, ref offset, windows1251, lengthBytes: 1));
    }

    /// <summary>
    /// A record longer than 8224 bytes is split, and the reader has to hand back the whole
    /// thing as though it had never been. This is the seam most legacy readers get wrong.
    /// </summary>
    [Fact]
    public void LongRecordsSurviveBeingSplitAndRejoined()
    {
        var payload = new byte[20_000];
        for (int i = 0; i < payload.Length; i++)
            payload[i] = (byte)(i % 251);

        using var writer = new BiffWriter(32_768);
        writer.WriteRecord(BiffRecordType.Sst, payload);

        using var reader = new BiffReader(writer.ToArray());
        Assert.True(reader.Read());
        Assert.Equal(BiffRecordType.Sst, reader.RecordType);
        Assert.True(reader.WasContinued);
        Assert.Equal(payload.Length, reader.Record.Length);
        Assert.True(payload.AsSpan().SequenceEqual(reader.Record.Span));

        // Three pieces: 8224, 8224 and the remainder.
        Assert.Equal(3, reader.ContinuationSizes.Count);
        Assert.False(reader.Read());
    }

    [Fact]
    public void ATruncatedRecordStopsTheReaderRatherThanReadingPastTheEnd()
    {
        // A header claiming 100 bytes over a stream that holds four.
        byte[] stream = [0x0A, 0x00, 0x64, 0x00, 0x01, 0x02, 0x03, 0x04];
        using var reader = new BiffReader(stream);
        Assert.False(reader.Read());
    }

    [Fact]
    public void CachedFormulaValuesDecodeByTag()
    {
        Span<byte> number = stackalloc byte[8];
        System.Buffers.Binary.BinaryPrimitives.WriteDoubleLittleEndian(number, 42.5);
        Assert.False(BiffValues.TryReadCachedValue(number, out CellValue value));
        Assert.Equal(42.5, value.GetNumber());

        Span<byte> boolean = [1, 0, 1, 0, 0, 0, 0xFF, 0xFF];
        Assert.False(BiffValues.TryReadCachedValue(boolean, out value));
        Assert.True(value.GetBoolean());

        Span<byte> error = [2, 0, 0x07, 0, 0, 0, 0xFF, 0xFF];
        Assert.False(BiffValues.TryReadCachedValue(error, out value));
        Assert.Equal(ExcelError.DivisionByZero, value.GetError());

        Span<byte> text = [0, 0, 0, 0, 0, 0, 0xFF, 0xFF];
        Assert.True(BiffValues.TryReadCachedValue(text, out _));
    }
}
