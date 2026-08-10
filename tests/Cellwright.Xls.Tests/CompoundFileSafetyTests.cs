using System.Buffers.Binary;
using Cellwright.Xls.Cfb;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>Hostile and uncommon MS-CFB layouts that must fail safely or read exactly.</summary>
public sealed class CompoundFileSafetyTests
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint FatSector = 0xFFFFFFFD;

    [Fact]
    public void ADirectoryChainThatCyclesAcrossTwoSectorsIsRejected()
    {
        byte[] file = Container(majorVersion: 3, sectorSize: 512, sectorCount: 3);
        Span<byte> fat = Sector(file, 512, 0);
        WriteUInt32(fat, 1, 2);
        WriteUInt32(fat, 2, 1);

        XlsFormatException failure = Assert.Throws<XlsFormatException>(() => CompoundFile.Parse(file));

        Assert.Contains("cyclic", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("directory", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void VersionFourPlacesSectorZeroAfterThe4096ByteHeaderSector()
    {
        byte[] file = Container(majorVersion: 4, sectorSize: 4096, sectorCount: 2);
        WriteRoot(Sector(file, 4096, 1));

        CompoundFile container = CompoundFile.Parse(file);

        Assert.Contains("Root Entry", container.EntryNames);
    }

    [Fact]
    public async Task VersionThreeIgnoresTheUninitializedUpperStreamSizeDword()
    {
        byte[] expected = new byte[5000];
        Random.Shared.NextBytes(expected);
        using var stream = new MemoryStream();
        await CompoundFileWriter.WriteAsync(
            stream, "Workbook", expected, TestContext.Current.CancellationToken);
        byte[] file = stream.ToArray();

        int directorySector = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(48)));
        int workbookEntry = ((directorySector + 1) * 512) + 128;
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(workbookEntry + 124), 0xDEADBEEF);

        byte[] actual = Assert.IsType<byte[]>(CompoundFile.Parse(file).TryReadStream("Workbook"));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task BinaryInputIsRejectedBeforeItExceedsTheConfiguredBudget()
    {
        using var input = new MemoryStream(new byte[2048], writable: false);
        var options = new XlsLoadOptions
        {
            LeaveOpen = true,
            Limits = new PackageLimits { MaxPartSize = 1024, MaxTotalSize = 1024 },
        };

        DocumentLoadLimitException failure = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await XlsWorkbook.LoadAsync(input, options, TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxInputBytes), failure.LimitName);
    }

    private static byte[] Container(ushort majorVersion, int sectorSize, int sectorCount)
    {
        var file = new byte[checked(sectorSize * (sectorCount + 1))];
        CompoundFile.Signature.CopyTo(file);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(24), 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(26), majorVersion);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(28), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(30), majorVersion == 3 ? (ushort)9 : (ushort)12);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(32), 6);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(40), majorVersion == 3 ? 0u : 1u);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(44), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(48), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(56), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(60), EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(68), EndOfChain);

        for (int i = 0; i < 109; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(76 + (i * 4)), FreeSector);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(76), 0);

        Span<byte> fat = Sector(file, sectorSize, 0);
        for (int offset = 0; offset < fat.Length; offset += 4)
            BinaryPrimitives.WriteUInt32LittleEndian(fat[offset..], FreeSector);
        WriteUInt32(fat, 0, FatSector);
        WriteUInt32(fat, 1, EndOfChain);
        return file;
    }

    private static void WriteRoot(Span<byte> directory)
    {
        const string name = "Root Entry\0";
        System.Text.Encoding.Unicode.GetBytes(name).CopyTo(directory);
        BinaryPrimitives.WriteUInt16LittleEndian(directory[64..], checked((ushort)(name.Length * 2)));
        directory[66] = 5;
        directory[67] = 1;
        BinaryPrimitives.WriteUInt32LittleEndian(directory[68..], FreeSector);
        BinaryPrimitives.WriteUInt32LittleEndian(directory[72..], FreeSector);
        BinaryPrimitives.WriteUInt32LittleEndian(directory[76..], FreeSector);
        BinaryPrimitives.WriteUInt32LittleEndian(directory[116..], EndOfChain);
    }

    private static Span<byte> Sector(byte[] file, int sectorSize, int number) =>
        file.AsSpan(checked((number + 1) * sectorSize), sectorSize);

    private static void WriteUInt32(Span<byte> sector, int index, uint value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(sector[(index * 4)..], value);
}
