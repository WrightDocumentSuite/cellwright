using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Cfb;

/// <summary>
/// Reads an OLE property set stream (MS-OLEPS 2.21), which is where a compound file keeps the
/// title, the author and the rest of what a modern package writes into <c>docProps</c>.
/// </summary>
/// <remarks>
/// The layout is a header naming one or two property sets by GUID, each of which is a table of
/// identifiers and offsets followed by the values themselves. Strings are encoded in whichever
/// code page the set declares as its first property, so that property has to be found before
/// any other can be decoded.
/// </remarks>
internal static class PropertySetStream
{
    /// <summary>The code page that means the strings are UTF-16 rather than bytes.</summary>
    private const int Unicode = 0x04B0;

    /// <summary>Property identifier of the code page every set states first (MS-OLEPS 2.18.2).</summary>
    private const uint CodePageProperty = 0x00000001;

    /// <summary>Property identifier of the name dictionary of a user-defined set (2.18.1).</summary>
    private const uint DictionaryProperty = 0x00000000;

    /// <summary>One property set: the values by identifier, and the names when it has them.</summary>
    /// <param name="Values">Property values by identifier.</param>
    /// <param name="Names">Names of user-defined properties by identifier, empty for a standard set.</param>
    public readonly record struct Section(
        IReadOnlyDictionary<uint, object> Values, IReadOnlyDictionary<uint, string> Names);

    /// <summary>Reads every property set in the stream, keyed by its format identifier.</summary>
    public static Dictionary<Guid, Section> Read(ReadOnlySpan<byte> stream)
    {
        var sets = new Dictionary<Guid, Section>();
        if (stream.Length < 48 || BinaryPrimitives.ReadUInt16LittleEndian(stream) != 0xFFFE)
            return sets;

        int count = (int)BinaryPrimitives.ReadUInt32LittleEndian(stream[24..]);
        if (count is < 1 or > 2)
            return sets;

        for (int i = 0; i < count; i++)
        {
            int entry = 28 + (i * 20);
            if (entry + 20 > stream.Length)
                break;

            var format = new Guid(stream.Slice(entry, 16));
            int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(stream[(entry + 16)..]);
            if (offset < 0 || offset >= stream.Length)
                continue;

            // A section runs to wherever the next one starts, or to the end of the stream. Its
            // own Size field says the same thing in principle, but files exist that understate
            // it — 42726.xls of the NPOI corpus declares 72 bytes and puts its author at 80 —
            // and believing that number loses every property past it.
            int end = stream.Length;
            if (i + 1 < count && entry + 40 <= stream.Length)
            {
                int next = (int)BinaryPrimitives.ReadUInt32LittleEndian(stream[(entry + 36)..]);
                if (next > offset && next <= stream.Length)
                    end = next;
            }

            sets[format] = ReadSection(stream[offset..end]);
        }

        return sets;
    }

    private static Section ReadSection(ReadOnlySpan<byte> section)
    {
        var values = new Dictionary<uint, object>();
        var names = new Dictionary<uint, string>();
        if (section.Length < 8)
            return new Section(values, names);

        int properties = (int)BinaryPrimitives.ReadUInt32LittleEndian(section[4..]);
        if (properties is < 0 or > 0xFFFF)
            return new Section(values, names);

        // The code page governs how every string in the set is decoded, and the table is not
        // required to list it first, so it is looked up before anything else is read.
        Encoding encoding = ResolveEncoding(section, properties);

        for (int i = 0; i < properties; i++)
        {
            int entry = 8 + (i * 8);
            if (entry + 8 > section.Length)
                break;

            uint id = BinaryPrimitives.ReadUInt32LittleEndian(section[entry..]);
            int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(section[(entry + 4)..]);
            if (offset < 0 || offset >= section.Length)
                continue;

            if (id == DictionaryProperty)
                ReadDictionary(section[offset..], encoding, names);
            else if (ReadValue(section[offset..], encoding) is { } value)
                values[id] = value;
        }

        return new Section(values, names);
    }

    private static Encoding ResolveEncoding(ReadOnlySpan<byte> section, int properties)
    {
        for (int i = 0; i < properties; i++)
        {
            int entry = 8 + (i * 8);
            if (entry + 8 > section.Length)
                break;

            if (BinaryPrimitives.ReadUInt32LittleEndian(section[entry..]) != CodePageProperty)
                continue;

            int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(section[(entry + 4)..]);
            if (offset < 0 || offset + 6 > section.Length)
                break;

            // The code page is written as a signed 16-bit value, so anything above 32767
            // arrives negative; 1200 for UTF-16 is the one that matters here.
            int codePage = BinaryPrimitives.ReadInt16LittleEndian(section[(offset + 4)..]) & 0xFFFF;
            return Resolve(codePage);
        }

        return Encoding.Latin1;
    }

    private static Encoding Resolve(int codePage)
    {
        if (codePage == Unicode)
            return Encoding.Unicode;

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (ArgumentException)
        {
            return Encoding.Latin1;
        }
        catch (NotSupportedException)
        {
            return Encoding.Latin1;
        }
    }

    /// <summary>
    /// Reads one <c>TypedPropertyValue</c> (MS-OLEPS 2.15): a type, two bytes of padding, and a
    /// value whose shape the type decides. Types this reader has no use for return null.
    /// </summary>
    private static object? ReadValue(ReadOnlySpan<byte> value, Encoding encoding)
    {
        if (value.Length < 4)
            return null;

        ushort type = BinaryPrimitives.ReadUInt16LittleEndian(value);
        ReadOnlySpan<byte> body = value[4..];
        return type switch
        {
            0x0002 when body.Length >= 2 => (int)BinaryPrimitives.ReadInt16LittleEndian(body),
            0x0003 when body.Length >= 4 => BinaryPrimitives.ReadInt32LittleEndian(body),
            0x0005 when body.Length >= 8 => BinaryPrimitives.ReadDoubleLittleEndian(body),
            0x000B when body.Length >= 2 => BinaryPrimitives.ReadInt16LittleEndian(body) != 0,
            0x001E => ReadCodePageString(body, encoding),
            0x001F => ReadCodePageString(body, Encoding.Unicode),
            0x0040 when body.Length >= 8 => ReadFileTime(body),
            _ => null,
        };
    }

    /// <summary>
    /// A <c>CodePageString</c> (2.5) or <c>UnicodeString</c> (2.7). Both are a byte count
    /// followed by characters that include a terminator, which is not part of the text.
    /// </summary>
    private static string? ReadCodePageString(ReadOnlySpan<byte> body, Encoding encoding)
    {
        if (body.Length < 4)
            return null;

        int size = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);

        // A Unicode string counts characters where a code-page string counts bytes.
        int bytes = encoding.Equals(Encoding.Unicode) ? size * 2 : size;
        if (bytes < 0 || 4 + bytes > body.Length)
            return null;

        string text = encoding.GetString(body.Slice(4, bytes));
        int nul = text.IndexOf('\0', StringComparison.Ordinal);
        return nul >= 0 ? text[..nul] : text;
    }

    /// <summary>A FILETIME, which counts hundred-nanosecond ticks from 1601 in UTC.</summary>
    private static DateTime? ReadFileTime(ReadOnlySpan<byte> body)
    {
        long ticks = BinaryPrimitives.ReadInt64LittleEndian(body);
        if (ticks <= 0)
            return null;

        try
        {
            return DateTime.FromFileTimeUtc(ticks);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    /// <summary>
    /// The dictionary of a user-defined set (MS-OLEPS 2.17): the names its properties go by,
    /// since the identifiers alone carry no meaning.
    /// </summary>
    private static void ReadDictionary(
        ReadOnlySpan<byte> body, Encoding encoding, Dictionary<uint, string> names)
    {
        if (body.Length < 4)
            return;

        int entries = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        if (entries is < 0 or > 0xFFFF)
            return;

        int offset = 4;
        bool unicode = encoding.Equals(Encoding.Unicode);
        for (int i = 0; i < entries && offset + 8 <= body.Length; i++)
        {
            uint id = BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
            int length = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 4)..]);
            offset += 8;

            int bytes = unicode ? length * 2 : length;
            if (bytes < 0 || offset + bytes > body.Length)
                return;

            string name = encoding.GetString(body.Slice(offset, bytes));
            int nul = name.IndexOf('\0', StringComparison.Ordinal);
            if (nul >= 0)
                name = name[..nul];
            if (name.Length > 0)
                names[id] = name;

            // A Unicode entry pads out to a multiple of four bytes; a byte one does not.
            offset += unicode ? (bytes + 3) / 4 * 4 : bytes;
        }
    }
}
