using System.Buffers.Binary;
using System.Text;
using Cellwright.Model;
using Cellwright.Xls.Reading;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// The shape of the HLINK record, one flag at a time (MS-OSHARED 2.3.7.1). Every field of the
/// structure is optional and they follow one another with nothing to resynchronise on, so these
/// are built by hand: a corpus file exercises one combination, and the combinations that break
/// a reader are the ones nobody happened to save.
/// </summary>
public sealed class HyperlinkRecordTests
{
    /// <summary>
    /// Bit 4 is what says the display text is present. Bit 2 says only that the author chose it
    /// rather than letting it default to the target. Requiring both means a link that set the
    /// first and not the second loses its text — and, far worse, reads its target out of the
    /// middle of that text, because every field after the missed one is off by its length.
    /// </summary>
    [Fact]
    public void DisplayTextIsPresentOnItsOwnFlag()
    {
        byte[] record = Build(
            flags: 0x0001 | 0x0002 | 0x0010,
            display: "Company site",
            target: "https://example.com/");

        Hyperlink link = ReadOne(record);

        Assert.Equal("Company site", link.Display);
        Assert.Equal("https://example.com/", link.Target);
    }

    /// <summary>The ordinary case, where both display bits are set, still reads the same.</summary>
    [Fact]
    public void DisplayTextIsPresentWhenBothFlagsAreSet()
    {
        byte[] record = Build(
            flags: 0x0001 | 0x0002 | 0x0004 | 0x0010,
            display: "Company site",
            target: "https://example.com/");

        Hyperlink link = ReadOne(record);

        Assert.Equal("Company site", link.Display);
        Assert.Equal("https://example.com/", link.Target);
    }

    /// <summary>
    /// A URL moniker may carry a serialisation block after the URL, and its length field counts
    /// those 24 bytes too. Decoding the whole of what the length covers leaves the block's bytes
    /// on the end of the address as unprintable characters.
    /// </summary>
    [Fact]
    public void AUrlKeepsNothingOfTheBlockThatFollowsIt()
    {
        byte[] record = Build(
            flags: 0x0001 | 0x0002,
            display: null,
            target: "https://example.com/",
            urlTrailer: true);

        Hyperlink link = ReadOne(record);

        Assert.Equal("https://example.com/", link.Target);
    }

    /// <summary>
    /// A moniker written as a plain string rather than as an OLE structure. The flag for it sits
    /// well above the ones a reader usually looks at, and missing it means reading the first
    /// sixteen bytes of the target as a class identifier and giving up on the link.
    /// </summary>
    [Fact]
    public void AMonikerSavedAsAStringIsStillATarget()
    {
        byte[] record = Build(
            flags: 0x0001 | 0x0002 | 0x0100,
            display: null,
            target: null,
            monikerString: "https://example.com/saved-as-string");

        Hyperlink link = ReadOne(record);

        Assert.Equal("https://example.com/saved-as-string", link.Target);
    }

    /// <summary>
    /// The tooltip is a record of its own that follows the link (MS-XLS 2.4.141), matched to it
    /// by the range it repeats.
    /// </summary>
    [Fact]
    public void TheTooltipComesFromTheRecordAfterTheLink()
    {
        Worksheet sheet = NewSheet();
        var sink = new WorksheetContentSink(sheet, styleMap: [], relationships: [], preserveMarkup: false, _ => { });
        HyperlinkRecord.Read(
            Build(flags: 0x0001 | 0x0002, display: null, target: "https://example.com/"), sink, _ => { });

        HyperlinkTooltipRecord.Read(BuildTooltip("Opens the company site"), sheet);

        Assert.Equal("Opens the company site", Assert.Single(sheet.Hyperlinks).Tooltip);
    }

    private static Worksheet NewSheet()
    {
        Workbook workbook = Workbook.Create();
        return workbook.Sheets.Add("Sheet1");
    }

    private static Hyperlink ReadOne(byte[] record)
    {
        Worksheet sheet = NewSheet();
        var sink = new WorksheetContentSink(sheet, styleMap: [], relationships: [], preserveMarkup: false, _ => { });
        HyperlinkRecord.Read(record, sink, _ => { });
        return Assert.Single(sheet.Hyperlinks);
    }

    /// <summary>Builds an HLINK body over A1 with the fields the flags ask for.</summary>
    private static byte[] Build(
        uint flags, string? display, string? target, bool urlTrailer = false, string? monikerString = null)
    {
        var bytes = new List<byte>();
        for (int i = 0; i < 4; i++)
            AppendUInt16(bytes, 0);                          // Ref8U over A1.

        bytes.AddRange(new byte[16]);                        // The hyperlink class id, unread.
        AppendUInt32(bytes, 2);                              // Stream version.
        AppendUInt32(bytes, flags);

        if (display is not null)
            AppendHyperlinkString(bytes, display);
        if (monikerString is not null)
            AppendHyperlinkString(bytes, monikerString);

        if (target is not null)
        {
            bytes.AddRange(new byte[]
            {
                0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11,
                0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
            });

            byte[] url = Encoding.Unicode.GetBytes(target + "\0");
            AppendUInt32(bytes, (uint)(url.Length + (urlTrailer ? 24 : 0)));
            bytes.AddRange(url);
            if (urlTrailer)
            {
                // The serialisation GUID the specification fixes, then its version and the URI
                // creation flags. Real bytes, because a block of zeros would decode to nothing
                // and let a reader that runs past the end of the URL look correct.
                bytes.AddRange(new byte[]
                {
                    0x79, 0x58, 0x81, 0xF4, 0x3B, 0x1D, 0x7F, 0x48,
                    0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63,
                });
                AppendUInt32(bytes, 0);
                AppendUInt32(bytes, 0x0000_0001);
            }
        }

        return [.. bytes];
    }

    private static byte[] BuildTooltip(string text)
    {
        var bytes = new List<byte>();
        AppendUInt16(bytes, 0x0800);                         // The record type, restated.
        for (int i = 0; i < 4; i++)
            AppendUInt16(bytes, 0);                          // Ref8U over A1.

        foreach (char c in text)
            AppendUInt16(bytes, c);
        AppendUInt16(bytes, 0);
        return [.. bytes];
    }

    private static void AppendHyperlinkString(List<byte> bytes, string text)
    {
        AppendUInt32(bytes, (uint)(text.Length + 1));
        bytes.AddRange(Encoding.Unicode.GetBytes(text + "\0"));
    }

    private static void AppendUInt16(List<byte> bytes, int value)
    {
        Span<byte> buffer = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16LittleEndian(buffer, (ushort)value);
        bytes.AddRange(buffer);
    }

    private static void AppendUInt32(List<byte> bytes, uint value)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
        bytes.AddRange(buffer);
    }
}
