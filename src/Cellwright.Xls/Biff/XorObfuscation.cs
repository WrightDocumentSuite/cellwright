namespace Cellwright.Xls.Biff;

/// <summary>
/// The XOR obfuscation of the earliest password-protected workbooks (MS-OFFCRYPTO 2.3.7). It is
/// not encryption in any modern sense — a sixteen-byte pad derived from the password, applied
/// over rotated bytes — but it is what stands between the reader and the content of these files.
/// </summary>
/// <remarks>
/// The two rotations here are placed differently from the specification and produce the same
/// bytes. It builds each pad entry as <c>ror1(character XOR key)</c> and then decrypts with
/// <c>ror5(data XOR pad)</c>; this builds <c>rol2(character XOR key)</c> and decrypts with
/// <c>ror5(data) XOR pad</c>. Rotation distributes over exclusive-or, and <c>ror5(ror1(v))</c>
/// is <c>rol2(v)</c> in eight bits, so the two agree — one rotation per byte instead of two.
/// </remarks>
internal static class XorObfuscation
{
    /// <summary>Builds the sixteen-byte pad a password produces.</summary>
    public static byte[] CreateArray(string password)
    {
        // Only the first fifteen characters take part; the rest of the pad is a fixed filler.
        string clipped = password.Length > 15 ? password[..15] : password;
        int key = CreateKey(clipped);

        var array = new byte[16];
        int index = clipped.Length;
        for (int i = 0; i < index; i++)
            array[i] = (byte)clipped[i];

        // The rest is the fixed filler, taken from its start however long the password was.
        for (int padIndex = 0; index < 16 && padIndex < PadArray.Length; padIndex++)
            array[index++] = PadArray[padIndex];

        Span<byte> keyBytes = [(byte)(key & 0xFF), (byte)((key >> 8) & 0xFF)];
        for (int i = 0; i < 16; i++)
        {
            array[i] ^= keyBytes[i & 1];
            array[i] = RotateLeft(array[i], 2);
        }

        return array;
    }

    /// <summary>Derives the sixteen-bit key a password hashes to.</summary>
    public static int CreateKey(string password)
    {
        if (password.Length == 0)
            return 0;

        string clipped = password.Length > 15 ? password[..15] : password;
        int key = InitialCode[clipped.Length - 1];
        int element = EncryptionMatrix.Length - 1;

        for (int i = clipped.Length - 1; i >= 0; i--)
        {
            int character = clipped[i];
            for (int bit = 0; bit < 7; bit++)
            {
                if ((character & 0x40) != 0)
                    key ^= EncryptionMatrix[element];
                character <<= 1;
                element--;
            }
        }

        return key & 0xFFFF;
    }

    /// <summary>The verifier a file stores so a reader can tell a wrong password from a right one.</summary>
    public static int CreateVerifier(string password)
    {
        if (password.Length == 0)
            return 0;

        string clipped = password.Length > 15 ? password[..15] : password;
        int verifier = 0;
        for (int i = clipped.Length - 1; i >= 0; i--)
        {
            verifier = ((verifier >> 14) & 0x01) | ((verifier << 1) & 0x7FFF);
            verifier ^= clipped[i];
        }

        verifier = ((verifier >> 14) & 0x01) | ((verifier << 1) & 0x7FFF);
        return (verifier ^ clipped.Length ^ 0xCE4B) & 0xFFFF;
    }

    /// <summary>
    /// Undoes the obfuscation over one record body. Each byte is rotated right by five and
    /// exclusive-ored with the pad, which advances one step per byte and wraps every sixteen.
    /// </summary>
    /// <param name="data">Bytes to transform in place.</param>
    /// <param name="pad">The sixteen-byte pad from <see cref="CreateArray"/>.</param>
    /// <param name="padIndex">Where in the pad to start; advanced past the data.</param>
    public static void Apply(Span<byte> data, ReadOnlySpan<byte> pad, ref int padIndex)
    {
        for (int i = 0; i < data.Length; i++)
        {
            byte value = RotateLeft(data[i], 3);
            data[i] = (byte)(value ^ pad[padIndex]);
            padIndex = (padIndex + 1) & 0x0F;
        }
    }

    private static byte RotateLeft(byte value, int count) =>
        (byte)(((value << count) | (value >> (8 - count))) & 0xFF);

    /// <summary>The filler that pads a short password out to sixteen bytes.</summary>
    private static ReadOnlySpan<byte> PadArray =>
    [
        0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80,
        0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
    ];

    /// <summary>Seed of the key, chosen by the password's length.</summary>
    private static ReadOnlySpan<int> InitialCode =>
    [
        0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E,
        0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3,
    ];

    /// <summary>
    /// The table the password's bits index into, fifteen rows of seven. It is a fixed part of
    /// the format with no structure worth deriving — reproduced exactly as specified.
    /// </summary>
    private static ReadOnlySpan<int> EncryptionMatrix =>
    [
        0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
        0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
        0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
        0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
        0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
        0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
        0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
        0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
        0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
        0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
        0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
        0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
        0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
        0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
        0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4,
    ];
}
