namespace Cellwright.Xls.Vba;

/// <summary>
/// The obfuscation the <c>PROJECT</c> stream wraps its protection fields in (MS-OVBA 2.4.3).
/// Every project Excel writes carries these fields whether or not anything is protected, so
/// their presence says nothing — only what they decode to does.
/// </summary>
/// <remarks>
/// This is not encryption and is not treated as such: the scheme is a rolling exclusive-or
/// keyed by a seed stored beside the data. It is decoded here for one purpose, telling a
/// protected project from an unprotected one, and the password itself is never reconstructed.
/// </remarks>
internal static class VbaProtection
{
    /// <summary>Whether a decoded <c>DPB</c> field means a password was set.</summary>
    /// <param name="hex">Value of the field, as the stream spells it.</param>
    /// <remarks>
    /// An unprotected project stores one byte of nothing (MS-OVBA 2.3.1.16); a protected one
    /// stores either the twenty-nine bytes of a hash and its key, or — the format allows it —
    /// the password itself as text. Anything but that single zero means a password. A field
    /// that will not decode says nothing either way.
    /// </remarks>
    public static bool IsPasswordProtected(ReadOnlySpan<char> hex)
    {
        byte[]? data = Decode(hex);
        return data is not null && (data.Length > 1 || (data.Length == 1 && data[0] != 0x00));
    }

    /// <summary>
    /// Reads the three protection bits out of a <c>CMG</c> field (MS-OVBA 2.3.1.15). Each says
    /// who asked for the project to be closed: the person editing it, the application hosting
    /// it, or the editor itself.
    /// </summary>
    public static (bool User, bool Host, bool Editor) DecodeProtectionState(ReadOnlySpan<char> hex)
    {
        if (Decode(hex) is not { Length: >= 1 } data)
            return default;

        byte flags = data[0];
        return ((flags & 0x01) != 0, (flags & 0x02) != 0, (flags & 0x04) != 0);
    }

    /// <summary>
    /// Whether a <c>GC</c> field says the project is visible (MS-OVBA 2.3.1.17). One byte:
    /// zero hides it, and the default when nothing says otherwise is visible.
    /// </summary>
    public static bool DecodeVisibility(ReadOnlySpan<char> hex) =>
        Decode(hex) is not { Length: >= 1 } data || data[0] != 0x00;

    /// <summary>Undoes the obfuscation, or returns <see langword="null"/> when the field is malformed.</summary>
    public static byte[]? Decode(ReadOnlySpan<char> hex)
    {
        if (hex.Length < 8 || hex.Length % 2 != 0)
            return null;

        Span<byte> raw = hex.Length <= 512 ? stackalloc byte[hex.Length / 2] : new byte[hex.Length / 2];
        if (!TryParseHex(hex, raw))
            return null;

        // The transform is seeded with the two header bytes in reverse: the project key is the
        // most recent encrypted byte, the version the one before it.
        byte seed = raw[0];
        byte encryptedPrevious = raw[2];
        byte encryptedBefore = raw[1];
        byte plainPrevious = (byte)(raw[2] ^ seed);
        if ((byte)(raw[1] ^ seed) != 2)
            return null;

        // The seed also says how many filler bytes sit between the header and the length.
        int position = 3;
        int ignored = (seed & 6) / 2;
        for (int i = 0; i < ignored && position < raw.Length; i++)
            Next(raw, ref position, ref encryptedPrevious, ref encryptedBefore, ref plainPrevious);

        int length = 0;
        for (int i = 0; i < 4; i++)
        {
            if (position >= raw.Length)
                return null;
            length |= Next(raw, ref position, ref encryptedPrevious, ref encryptedBefore, ref plainPrevious) << (i * 8);
        }

        if (length < 0 || length > raw.Length - position)
            return null;

        var data = new byte[length];
        for (int i = 0; i < length; i++)
            data[i] = (byte)Next(raw, ref position, ref encryptedPrevious, ref encryptedBefore, ref plainPrevious);
        return data;
    }

    /// <summary>
    /// One step of the rolling transform: each byte is undone with the sum of the byte before
    /// last as it was stored and the previous byte as it came out.
    /// </summary>
    private static int Next(
        ReadOnlySpan<byte> raw, ref int position, ref byte encryptedPrevious, ref byte encryptedBefore, ref byte plainPrevious)
    {
        byte encrypted = raw[position++];
        byte plain = (byte)(encrypted ^ (byte)(encryptedBefore + plainPrevious));
        encryptedBefore = encryptedPrevious;
        encryptedPrevious = encrypted;
        plainPrevious = plain;
        return plain;
    }

    private static bool TryParseHex(scoped ReadOnlySpan<char> hex, scoped Span<byte> destination)
    {
        for (int i = 0; i < destination.Length; i++)
        {
            if (!byte.TryParse(hex.Slice(i * 2, 2), System.Globalization.NumberStyles.HexNumber, null, out destination[i]))
                return false;
        }

        return true;
    }
}
