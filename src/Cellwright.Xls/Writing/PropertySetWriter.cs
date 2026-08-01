using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the two OLE property set streams that carry a document's title, author and the rest
/// (MS-OLEPS 2.21), so that a workbook saved as .xls keeps what it had as .xlsx.
/// </summary>
/// <remarks>
/// Strings go out as UTF-16 throughout, with the set's code page declared as CP_WINUNICODE.
/// The alternative is choosing a byte code page per document and hoping every name in it fits,
/// which is how an author called Отдел финансов turns into question marks.
/// </remarks>
internal static class PropertySetWriter
{
    private static readonly Guid SummaryInformation = new("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
    private static readonly Guid DocumentSummaryInformation = new("D5CDD502-2E9C-101B-9397-08002B2CF9AE");

    /// <summary>Name of the stream carrying the summary set.</summary>
    public const string SummaryStreamName = "\u0005SummaryInformation";

    /// <summary>Name of the stream carrying the document summary set.</summary>
    public const string DocumentSummaryStreamName = "\u0005DocumentSummaryInformation";

    private const ushort CodePageUnicode = 0x04B0;
    private const ushort TypeI2 = 0x0002;
    private const ushort TypeLpwstr = 0x001F;
    private const ushort TypeFileTime = 0x0040;

    /// <summary>
    /// Builds the summary stream, or returns <see langword="null"/> when there is nothing in it
    /// worth writing. A property set holding only its own code page is noise in the file.
    /// </summary>
    public static byte[]? WriteSummary(WorkbookProperties properties)
    {
        var values = new List<(uint Id, object Value)>();
        Add(values, 0x02, properties.Title);
        Add(values, 0x03, properties.Subject);
        Add(values, 0x04, properties.Author);
        Add(values, 0x05, properties.Keywords);
        Add(values, 0x06, properties.Description);
        Add(values, 0x08, properties.LastModifiedBy);
        Add(values, 0x12, properties.Application);
        Add(values, 0x0C, properties.Created);
        Add(values, 0x0D, properties.Modified);

        return values.Count == 0 ? null : Build(SummaryInformation, values);
    }

    /// <summary>Builds the document summary stream, or null when it would say nothing.</summary>
    public static byte[]? WriteDocumentSummary(WorkbookProperties properties)
    {
        var values = new List<(uint Id, object Value)>();
        Add(values, 0x02, properties.Category);
        Add(values, 0x0E, properties.Manager);
        Add(values, 0x0F, properties.Company);
        Add(values, 0x1B, properties.ContentStatus);

        return values.Count == 0 ? null : Build(DocumentSummaryInformation, values);
    }

    private static void Add(List<(uint, object)> values, uint id, string? text)
    {
        if (!string.IsNullOrEmpty(text))
            values.Add((id, text));
    }

    private static void Add(List<(uint, object)> values, uint id, DateTime? moment)
    {
        if (moment is { } value)
            values.Add((id, value));
    }

    /// <summary>
    /// Lays out one property set stream: the header naming the set, then a table of identifiers
    /// and offsets, then the values. Every offset has to be a multiple of four.
    /// </summary>
    private static byte[] Build(Guid format, List<(uint Id, object Value)> values)
    {
        // The code page comes first so that a reader meets it before any string it governs.
        var properties = new List<(uint Id, byte[] Bytes)>(values.Count + 1)
        {
            (0x00000001, CodePage()),
        };

        foreach ((uint id, object value) in values)
            properties.Add((id, Encode(value)));

        int tableSize = 8 + (properties.Count * 8);
        int sectionSize = tableSize + properties.Sum(static property => property.Bytes.Length);
        const int SectionOffset = 48;

        var bytes = new byte[SectionOffset + sectionSize];
        Span<byte> span = bytes;

        BinaryPrimitives.WriteUInt16LittleEndian(span, 0xFFFE);              // Byte order.
        BinaryPrimitives.WriteUInt16LittleEndian(span[2..], 0);              // Version 0.
        BinaryPrimitives.WriteUInt32LittleEndian(span[4..], 0x00020105);     // System identifier, ignored.
        BinaryPrimitives.WriteUInt32LittleEndian(span[24..], 1);             // One property set.
        format.TryWriteBytes(span[28..44]);
        BinaryPrimitives.WriteUInt32LittleEndian(span[44..], SectionOffset);

        Span<byte> section = span[SectionOffset..];
        BinaryPrimitives.WriteUInt32LittleEndian(section, (uint)sectionSize);
        BinaryPrimitives.WriteUInt32LittleEndian(section[4..], (uint)properties.Count);

        int offset = tableSize;
        for (int i = 0; i < properties.Count; i++)
        {
            int entry = 8 + (i * 8);
            BinaryPrimitives.WriteUInt32LittleEndian(section[entry..], properties[i].Id);
            BinaryPrimitives.WriteUInt32LittleEndian(section[(entry + 4)..], (uint)offset);
            properties[i].Bytes.CopyTo(section[offset..]);
            offset += properties[i].Bytes.Length;
        }

        return bytes;
    }

    private static byte[] CodePage()
    {
        var bytes = new byte[8];
        BinaryPrimitives.WriteUInt16LittleEndian(bytes, TypeI2);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(4), CodePageUnicode);
        return bytes;
    }

    private static byte[] Encode(object value) => value switch
    {
        string text => EncodeString(text),
        DateTime moment => EncodeFileTime(moment),
        _ => throw new InvalidOperationException($"No property encoding for {value.GetType()}."),
    };

    /// <summary>
    /// A UnicodeString property: the type, a count of characters including the terminator, and
    /// the characters, padded out so that the next property still starts on a multiple of four.
    /// </summary>
    private static byte[] EncodeString(string text)
    {
        int characters = text.Length + 1;
        int payload = characters * 2;
        var bytes = new byte[8 + RoundUpToMultiple(payload, 4)];

        BinaryPrimitives.WriteUInt16LittleEndian(bytes, TypeLpwstr);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4), (uint)characters);
        Encoding.Unicode.GetBytes(text, bytes.AsSpan(8));
        return bytes;
    }

    private static byte[] EncodeFileTime(DateTime moment)
    {
        var bytes = new byte[12];
        BinaryPrimitives.WriteUInt16LittleEndian(bytes, TypeFileTime);

        DateTime utc = moment.Kind == DateTimeKind.Unspecified
            ? DateTime.SpecifyKind(moment, DateTimeKind.Utc)
            : moment.ToUniversalTime();
        BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(4), utc.ToFileTimeUtc());
        return bytes;
    }

    private static int RoundUpToMultiple(int value, int multiple) => (value + multiple - 1) / multiple * multiple;
}
