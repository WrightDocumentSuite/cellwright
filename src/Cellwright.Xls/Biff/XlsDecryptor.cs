using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace Cellwright.Xls.Biff;

/// <summary>
/// Undoes the "password to open" protection of a legacy workbook: the XOR obfuscation of the
/// earliest files, and the two RC4 schemes that followed.
/// </summary>
/// <remarks>
/// None of these were ever strong, and the default password <c>VelvetSweatshop</c> — which
/// Excel applies without telling anyone, so that a workbook can be encrypted and still open on
/// a double click — is tried first. The RC4 schemes run over the stream in 1024-byte blocks,
/// each with a key of its own; a handful of records are left in the clear because a reader has
/// to find its bearings before it can decrypt anything.
/// </remarks>
internal sealed class XlsDecryptor : IRecordDecryptor
{
    /// <summary>The password Excel uses when a file is encrypted but not password-protected.</summary>
    public const string DefaultPassword = "VelvetSweatshop";

    private const int BlockSize = 1024;

    private readonly Scheme _scheme;
    private readonly byte[] _key;
    private readonly byte[] _xorPad;
    private readonly int _keyBytes;
    private readonly byte[] _buffer = new byte[BiffReader.MaxRecordSize];

    private XlsDecryptor(Scheme scheme, byte[] key, byte[] xorPad, int keyBytes = 0)
    {
        _scheme = scheme;
        _key = key;
        _xorPad = xorPad;
        _keyBytes = keyBytes;
    }

    private enum Scheme
    {
        /// <summary>The pad-and-rotate obfuscation of Excel 95 and earlier.</summary>
        Xor,

        /// <summary>RC4 with a 40-bit key derived through MD5, as Excel 97 wrote.</summary>
        Rc4,

        /// <summary>RC4 with a key derived through SHA-1 by the Windows crypto provider.</summary>
        Rc4CryptoApi,
    }

    /// <summary>Builds a decryptor from a <c>FilePass</c> record, or reports why it cannot.</summary>
    /// <param name="filePass">Body of the FilePass record.</param>
    /// <param name="password">Password to try; the default one is used when none is given.</param>
    /// <param name="decryptor">The decryptor on success.</param>
    /// <param name="failure">Why it failed, when it did.</param>
    public static bool TryCreate(
        ReadOnlySpan<byte> filePass, string? password, out XlsDecryptor? decryptor, out string? failure)
    {
        decryptor = null;
        failure = null;
        if (filePass.Length < 4)
        {
            failure = "The FilePass record is too short to say how the file is encrypted.";
            return false;
        }

        // Before BIFF8 the record has no type field: four bytes of key and verifier is all it
        // is, and the scheme is always the XOR obfuscation.
        if (filePass.Length == 4)
            return TryCreateXor(filePass, password, out decryptor, out failure);

        ushort type = BinaryPrimitives.ReadUInt16LittleEndian(filePass);
        if (type == 0)
            return TryCreateXor(filePass[2..], password, out decryptor, out failure);

        if (filePass.Length < 6)
        {
            failure = "The FilePass record declares RC4 but carries no encryption header.";
            return false;
        }

        ushort major = BinaryPrimitives.ReadUInt16LittleEndian(filePass[2..]);
        ushort minor = BinaryPrimitives.ReadUInt16LittleEndian(filePass[4..]);

        // Major 1 is the original 40-bit RC4; 2, 3 and 4 are the CryptoAPI form, which carries a
        // full encryption header describing the algorithm and key length.
        return major == 1 && minor == 1
            ? TryCreateRc4(filePass[6..], password, out decryptor, out failure)
            : TryCreateRc4CryptoApi(filePass[6..], password, out decryptor, out failure);
    }

    private static bool TryCreateXor(
        ReadOnlySpan<byte> body, string? password, out XlsDecryptor? decryptor, out string? failure)
    {
        decryptor = null;
        failure = null;
        if (body.Length < 4)
        {
            failure = "The XOR-obfuscated file carries no key.";
            return false;
        }

        int key = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int verifier = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);

        foreach (string candidate in Candidates(password))
        {
            if (XorObfuscation.CreateKey(candidate) != key || XorObfuscation.CreateVerifier(candidate) != verifier)
                continue;

            decryptor = new XlsDecryptor(Scheme.Xor, [], XorObfuscation.CreateArray(candidate));
            return true;
        }

        failure = password is null
            ? "The workbook is encrypted and needs a password."
            : "The password does not open this workbook.";
        return false;
    }

    private static bool TryCreateRc4(
        ReadOnlySpan<byte> header, string? password, out XlsDecryptor? decryptor, out string? failure)
    {
        decryptor = null;
        failure = null;
        if (header.Length < 48)
        {
            failure = "The RC4 encryption header is truncated.";
            return false;
        }

        ReadOnlySpan<byte> salt = header[..16];
        ReadOnlySpan<byte> verifier = header[16..32];
        ReadOnlySpan<byte> verifierHash = header[32..48];

        foreach (string candidate in Candidates(password))
        {
            byte[] key = Rc4KeyDigest(candidate, salt);
            if (!VerifiesRc4(key, verifier, verifierHash))
                continue;

            decryptor = new XlsDecryptor(Scheme.Rc4, key, []);
            return true;
        }

        failure = password is null
            ? "The workbook is encrypted and needs a password."
            : "The password does not open this workbook.";
        return false;
    }

    private static bool TryCreateRc4CryptoApi(
        ReadOnlySpan<byte> header, string? password, out XlsDecryptor? decryptor, out string? failure)
    {
        decryptor = null;
        failure = null;
        if (header.Length < 12)
        {
            failure = "The CryptoAPI encryption header is truncated.";
            return false;
        }

        int headerSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(header[4..]);
        int verifierOffset = 8 + headerSize;
        if (headerSize < 32 || verifierOffset + 40 > header.Length)
        {
            failure = "The CryptoAPI encryption header does not fit the record.";
            return false;
        }

        int keyBits = (int)BinaryPrimitives.ReadUInt32LittleEndian(header[(8 + 16)..]);

        // A key size of zero means the provider's default, which for RC4 is 40 bits.
        int keyBytes = keyBits == 0 ? 5 : Math.Clamp(keyBits / 8, 5, 16);

        ReadOnlySpan<byte> verifierBlock = header[verifierOffset..];
        int saltSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(verifierBlock);
        if (saltSize != 16 || verifierBlock.Length < 4 + 16 + 16 + 4 + 20)
        {
            failure = "The CryptoAPI verifier block is not the shape the specification describes.";
            return false;
        }

        ReadOnlySpan<byte> salt = verifierBlock.Slice(4, 16);
        ReadOnlySpan<byte> verifier = verifierBlock.Slice(20, 16);
        ReadOnlySpan<byte> verifierHash = verifierBlock.Slice(40, 20);

        foreach (string candidate in Candidates(password))
        {
            byte[] key = CryptoApiSecretKey(candidate, salt);
            if (!VerifiesCryptoApi(key, keyBytes, verifier, verifierHash))
                continue;

            decryptor = new XlsDecryptor(Scheme.Rc4CryptoApi, key, [], keyBytes);
            return true;
        }

        failure = password is null
            ? "The workbook is encrypted and needs a password."
            : "The password does not open this workbook.";
        return false;
    }

    /// <summary>The passwords to try, default first, so that a silently encrypted file just opens.</summary>
    private static IEnumerable<string> Candidates(string? password)
    {
        yield return DefaultPassword;
        if (!string.IsNullOrEmpty(password) && password != DefaultPassword)
            yield return password;
    }

    /// <inheritdoc/>
    public ReadOnlyMemory<byte> Decrypt(byte[] stream, int offset, int size, ushort recordType)
    {
        if (size == 0)
            return ReadOnlyMemory<byte>.Empty;

        // A reader must be able to find the file's version and its encryption header before it
        // has a key, so those records are never encrypted (MS-XLS 2.2.10).
        if (IsPlain(recordType))
            return stream.AsMemory(offset, size);

        Span<byte> destination = _buffer.AsSpan(0, size);
        stream.AsSpan(offset, size).CopyTo(destination);

        if (_scheme == Scheme.Xor)
        {
            // The pad's starting position depends on where the record ends, which is what makes
            // this transformable record by record rather than only front to back.
            int padIndex = (offset + size) & 0x0F;
            XorObfuscation.Apply(destination, _xorPad, ref padIndex);
        }
        else
        {
            DecryptRc4Blocks(destination, offset);
        }

        // The sheet offset that opens a BoundSheet8 is left in the clear, so that a reader can
        // find the sheets before it has a key. The cipher still runs over those four bytes —
        // they count towards its position — but the plain ones are what the record means.
        if (recordType == BiffRecordType.BoundSheet8 && size >= 4)
            stream.AsSpan(offset, 4).CopyTo(destination);

        return _buffer.AsMemory(0, size);
    }

    private static bool IsPlain(ushort recordType) => recordType is
        BiffRecordType.Bof or
        BiffRecordType.FilePass or
        BiffRecordType.Interfacehdr or
        0x002A or  // WriteProt.
        0x00E1 or  // InterfaceHdr, spelled out for the reader's benefit.
        0x0195;    // FileSharing, which carries the write-reservation password.

    /// <summary>
    /// Runs RC4 over one record. The keystream restarts every 1024 bytes of the stream with a
    /// block-numbered key, so the cipher has to be positioned at the record's offset first.
    /// </summary>
    private void DecryptRc4Blocks(Span<byte> destination, int offset)
    {
        int position = offset;
        int written = 0;
        while (written < destination.Length)
        {
            int block = position / BlockSize;
            int withinBlock = position % BlockSize;
            int take = Math.Min(BlockSize - withinBlock, destination.Length - written);

            var cipher = new Rc4(BlockKey(block));
            cipher.Skip(withinBlock);
            cipher.Apply(destination.Slice(written, take));

            written += take;
            position += take;
        }
    }

    private byte[] BlockKey(int block) =>
        _scheme == Scheme.Rc4 ? Rc4BlockKey(_key, block) : CryptoApiBlockKey(_key, _keyBytes, block);

    /// <summary>
    /// Derives the five-byte digest the original RC4 scheme hangs everything off: the password
    /// hashed, then its first five bytes stirred with the salt sixteen times over
    /// (MS-OFFCRYPTO 2.3.6.2).
    /// </summary>
    private static byte[] Rc4KeyDigest(string password, ReadOnlySpan<byte> salt)
    {
        string clipped = password.Length > 16 ? password[..16] : password;
        byte[] encoded = Encoding.Unicode.GetBytes(clipped);

        Span<byte> passwordHash = stackalloc byte[16];
        MD5.HashData(encoded, passwordHash);

        Span<byte> intermediate = stackalloc byte[16 * 21];
        for (int i = 0; i < 16; i++)
        {
            passwordHash[..5].CopyTo(intermediate[(i * 21)..]);
            salt.CopyTo(intermediate[((i * 21) + 5)..]);
        }

        Span<byte> digest = stackalloc byte[16];
        MD5.HashData(intermediate, digest);
        return digest[..5].ToArray();
    }

    /// <summary>
    /// The RC4 key of one block: the five-byte digest with the block number appended, hashed.
    /// All sixteen bytes of that hash are the key — truncating it to the nominal 40 bits is the
    /// classic way to get this wrong.
    /// </summary>
    private static byte[] Rc4BlockKey(byte[] keyDigest, int block)
    {
        Span<byte> material = stackalloc byte[keyDigest.Length + 4];
        keyDigest.CopyTo(material);
        BinaryPrimitives.WriteInt32LittleEndian(material[keyDigest.Length..], block);

        Span<byte> digest = stackalloc byte[16];
        MD5.HashData(material, digest);
        return digest.ToArray();
    }

    private static bool VerifiesRc4(byte[] keyDigest, ReadOnlySpan<byte> verifier, ReadOnlySpan<byte> verifierHash)
    {
        var cipher = new Rc4(Rc4BlockKey(keyDigest, 0));
        Span<byte> plainVerifier = stackalloc byte[16];
        Span<byte> plainHash = stackalloc byte[16];
        verifier.CopyTo(plainVerifier);
        verifierHash.CopyTo(plainHash);

        // One cipher decrypts both in turn: the keystream runs on from the first into the second.
        cipher.Apply(plainVerifier);
        cipher.Apply(plainHash);

        Span<byte> expected = stackalloc byte[16];
        MD5.HashData(plainVerifier, expected);
        return expected.SequenceEqual(plainHash);
    }

    /// <summary>The CryptoAPI base key: the whole hash of the salt and the password.</summary>
    private static byte[] CryptoApiSecretKey(string password, ReadOnlySpan<byte> salt)
    {
        byte[] encoded = Encoding.Unicode.GetBytes(password);
        Span<byte> material = stackalloc byte[salt.Length + encoded.Length];
        salt.CopyTo(material);
        encoded.CopyTo(material[salt.Length..]);

        Span<byte> digest = stackalloc byte[20];
        SHA1.HashData(material, digest);
        return digest.ToArray();
    }

    /// <summary>
    /// The CryptoAPI key of one block: the base key with the block number, hashed and cut to
    /// the declared length. A 40-bit key is then padded back out to 128 with zeros
    /// (MS-OFFCRYPTO 2.3.5.2) — the padding is part of the key, and RC4 keyed with five bytes
    /// produces an entirely different keystream from RC4 keyed with those five and eleven zeros.
    /// </summary>
    private static byte[] CryptoApiBlockKey(byte[] secretKey, int keyBytes, int block)
    {
        Span<byte> material = stackalloc byte[secretKey.Length + 4];
        secretKey.CopyTo(material);
        BinaryPrimitives.WriteInt32LittleEndian(material[secretKey.Length..], block);

        Span<byte> digest = stackalloc byte[20];
        SHA1.HashData(material, digest);

        if (keyBytes != 5)
            return digest[..keyBytes].ToArray();

        var padded = new byte[16];
        digest[..5].CopyTo(padded);
        return padded;
    }

    private static bool VerifiesCryptoApi(
        byte[] secretKey, int keyBytes, ReadOnlySpan<byte> verifier, ReadOnlySpan<byte> verifierHash)
    {
        var cipher = new Rc4(CryptoApiBlockKey(secretKey, keyBytes, 0));
        Span<byte> plainVerifier = stackalloc byte[16];
        Span<byte> plainHash = stackalloc byte[20];
        verifier.CopyTo(plainVerifier);
        verifierHash.CopyTo(plainHash);
        cipher.Apply(plainVerifier);
        cipher.Apply(plainHash);

        Span<byte> expected = stackalloc byte[20];
        SHA1.HashData(plainVerifier, expected);
        return expected.SequenceEqual(plainHash);
    }
}
