using Cellwright.Xls.Xlsb;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The record framing of a binary workbook: a type and a size, each seven bits at a time. The
/// reader is the oracle — anything the writer emits has to come back through it unchanged.
/// </summary>
public sealed class Biff12WriterTests
{
    private static Biff12Reader ReaderOver(Biff12Writer writer) => new(writer.ToArray());

    [Theory]
    [InlineData(0)]        // One byte of type.
    [InlineData(127)]      // The last one-byte type.
    [InlineData(128)]      // The first that needs two.
    [InlineData(1024)]
    [InlineData(16383)]    // The widest the two-byte form holds.
    public void ARecordTypeComesBackWhateverItsWidth(int type)
    {
        using var writer = new Biff12Writer();
        writer.WriteRecord(type, [1, 2, 3]);

        Biff12Reader reader = ReaderOver(writer);

        Assert.True(reader.Read());
        Assert.Equal(type, reader.RecordType);
        Assert.Equal([1, 2, 3], reader.Record.ToArray());
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(127)]      // One byte of size.
    [InlineData(128)]      // Two.
    [InlineData(16383)]
    [InlineData(16384)]    // Three.
    [InlineData(70000)]    // Past what the older format could hold in one record at all.
    public void ABodyComesBackWhateverItsLength(int length)
    {
        var body = new byte[length];
        for (int i = 0; i < length; i++)
            body[i] = (byte)(i % 251);

        using var writer = new Biff12Writer();
        writer.WriteRecord(Biff12RecordType.CellReal, body);

        Biff12Reader reader = ReaderOver(writer);

        Assert.True(reader.Read());
        Assert.Equal(length, reader.Record.Length);
        Assert.Equal(body, reader.Record.ToArray());
    }

    [Fact]
    public void RecordsFollowOneAnother()
    {
        using var writer = new Biff12Writer();
        writer.WriteRecord(Biff12RecordType.BeginSheetData);
        writer.WriteRecord(Biff12RecordType.CellBlank, [7]);
        writer.WriteRecord(Biff12RecordType.EndSheetData);

        Biff12Reader reader = ReaderOver(writer);

        Assert.True(reader.Read());
        Assert.Equal(Biff12RecordType.BeginSheetData, reader.RecordType);
        Assert.True(reader.Read());
        Assert.Equal(Biff12RecordType.CellBlank, reader.RecordType);
        Assert.Equal(1, reader.Record.Length);
        Assert.True(reader.Read());
        Assert.Equal(Biff12RecordType.EndSheetData, reader.RecordType);
        Assert.False(reader.Read());
    }

    [Fact]
    public void AStringIsCountedInCharactersAndWrittenInTwoBytesEach()
    {
        using var body = new Biff12Writer();
        body.AppendString("Ok");

        // Four bytes of count, then two per character.
        Assert.Equal(8, body.Length);

        byte[] bytes = body.ToArray();
        Assert.Equal(2u, System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bytes));
        Assert.Equal("Ok", System.Text.Encoding.Unicode.GetString(bytes, 4, 4));
    }

    [Fact]
    public void ANullStringIsNotAnEmptyOne()
    {
        using var absent = new Biff12Writer();
        absent.AppendNullableString(null);

        using var empty = new Biff12Writer();
        empty.AppendNullableString(string.Empty);

        Assert.Equal(4, absent.Length);
        Assert.Equal(0xFFFFFFFF, System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(absent.ToArray()));
        Assert.Equal(4, empty.Length);
        Assert.Equal(0u, System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(empty.ToArray()));
    }

    [Fact]
    public void NonLatinTextSurvives()
    {
        const string Text = "Ведомость 会計";

        using var body = new Biff12Writer();
        body.AppendString(Text);

        byte[] bytes = body.ToArray();
        Assert.Equal((uint)Text.Length, System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bytes));
        Assert.Equal(Text, System.Text.Encoding.Unicode.GetString(bytes, 4, Text.Length * 2));
    }

    [Fact]
    public void ResettingKeepsTheBufferAndDropsTheBytes()
    {
        using var writer = new Biff12Writer();
        writer.AppendUInt32(7);
        Assert.Equal(4, writer.Length);

        writer.Reset();

        Assert.Equal(0, writer.Length);
        writer.AppendByte(1);
        Assert.Equal(1, writer.Length);
    }
}
