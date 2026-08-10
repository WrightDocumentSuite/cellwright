using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Xls.Crypto;

/// <summary>
/// The encryption Office 2010 introduced and everything since writes: a block cipher in CBC
/// mode, a key stretched through a hundred thousand hash rounds, and — unlike everything before
/// it — an XML descriptor in the file saying which algorithms were used rather than fixing them
/// in the format. The package is encrypted in 4096-byte segments, each with an initialization
/// vector of its own, so that a consumer can seek without decrypting what came before.
/// </summary>
internal static class AgileEncryption
{
    /// <summary>
    /// Largest Agile EncryptionInfo XML descriptor. Real password and certificate descriptors
    /// are tiny; this bound prevents XmlReader from first materializing a hostile huge attribute.
    /// </summary>
    internal const int MaxDescriptorBytes = 1 * 1024 * 1024;

    /// <summary>Segment size the specification fixes for the encrypted package (MS-OFFCRYPTO 2.3.4.15).</summary>
    private const int SegmentSize = 4096;

    /// <summary>What a key or an initialization vector is padded out with (MS-OFFCRYPTO 2.3.4.11).</summary>
    private const byte KeyPadding = 0x36;

    /// <summary>Salt length written, in bytes. Excel writes sixteen, and so does this.</summary>
    private const int SaltBytes = 16;

    /// <summary>Rounds the password is stretched by when writing; the maximum Office uses.</summary>
    private const int SpinCount = 100_000;

    /// <summary>
    /// Largest encoded salt, verifier, key or HMAC value accepted in EncryptionInfo. Real values
    /// are at most a few hundred characters; this leaves ample extension room while preventing a
    /// hostile XML attribute from causing another proportional decoded-buffer allocation.
    /// </summary>
    private const int MaxEncodedCryptoValueCharacters = 4 * 1024;

    // Block keys that separate the purposes the one key is used for (2.3.4.13 and 2.3.4.14).
    private static ReadOnlySpan<byte> VerifierInputBlock => [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];

    private static ReadOnlySpan<byte> VerifierValueBlock => [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];

    private static ReadOnlySpan<byte> KeyValueBlock => [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

    private static ReadOnlySpan<byte> HmacKeyBlock => [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];

    private static ReadOnlySpan<byte> HmacValueBlock => [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

    /// <summary>
    /// Recovers the package from its encrypted stream, or returns <see langword="null"/> when
    /// none of the passwords opens it.
    /// </summary>
    /// <param name="descriptor">The XML that follows the version header of <c>EncryptionInfo</c>.</param>
    /// <param name="encryptedPackage">Content of the <c>EncryptedPackage</c> stream.</param>
    /// <param name="passwords">Passwords to try, in order.</param>
    /// <param name="warn">Reports a package whose integrity check does not hold.</param>
    /// <param name="cancellationToken">Cancels password stretching and package decryption.</param>
    /// <param name="loadBudget">Shared XML and expansion counters of the outer document.</param>
    /// <exception cref="XlsFormatException">The descriptor describes something unreadable.</exception>
    public static byte[]? Decrypt(
        ReadOnlySpan<byte> descriptor, byte[] encryptedPackage, IEnumerable<string> passwords,
        Action<WorkbookWarning>? warn = null,
        CancellationToken cancellationToken = default,
        DocumentLoadBudgetState? loadBudget = null)
    {
        Descriptor parsed = Parse(descriptor, loadBudget, cancellationToken);
        foreach (string password in passwords)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryUnlock(parsed, password, cancellationToken) is not { } secretKey)
                continue;

            VerifyIntegrity(parsed, secretKey, encryptedPackage, warn);
            return PackageBytes(parsed.KeyData, secretKey, encryptedPackage, cancellationToken);
        }

        return null;
    }

    /// <summary>The two streams an encrypted package is written as.</summary>
    /// <param name="Descriptor">The XML that goes after the version header of <c>EncryptionInfo</c>.</param>
    /// <param name="EncryptedPackage">Content of the <c>EncryptedPackage</c> stream.</param>
    public readonly record struct Written(byte[] Descriptor, byte[] EncryptedPackage);

    /// <summary>
    /// Encrypts a package under a password, producing what <see cref="Decrypt"/> reads back.
    /// </summary>
    /// <remarks>
    /// Only one of the many shapes this format allows is written: AES-256 in CBC mode with
    /// SHA-512 and a hundred thousand rounds, which is what Excel itself writes today. Reading
    /// has to accept whatever a file happens to say; writing has no such obligation, and
    /// offering a choice of weaker ciphers would be offering a way to get it wrong.
    /// </remarks>
    /// <param name="package">The plain package to encrypt.</param>
    /// <param name="password">Password that will open it.</param>
    public static Written Encrypt(ReadOnlySpan<byte> package, string password)
    {
        var keyData = new Cipher(
            CryptoHash.Sha512, CipherKind.Aes, KeyBytes: 32, BlockBytes: 16, Cfb: false,
            Salt: RandomNumberGenerator.GetBytes(SaltBytes));
        var passwordKey = keyData with { Salt = RandomNumberGenerator.GetBytes(SaltBytes) };

        // The password does not encrypt the content: it wraps a key drawn at random, which is
        // what would let one file be opened by several credentials.
        byte[] secretKey = RandomNumberGenerator.GetBytes(keyData.KeyBytes);
        byte[] encryptedPackage = EncryptPackage(keyData, secretKey, package);

        (byte[] hmacKey, byte[] hmacValue) = SealIntegrity(keyData, secretKey, encryptedPackage);
        Locked locked = Lock(passwordKey, password, secretKey);

        return new Written(
            Describe(keyData, passwordKey, locked, hmacKey, hmacValue),
            encryptedPackage);
    }

    /// <summary>Everything the password encryptor holds, once the password has been applied.</summary>
    private readonly record struct Locked(byte[] VerifierInput, byte[] VerifierValue, byte[] EncryptedKey);

    /// <summary>
    /// Wraps the secret key with the password, and builds the pair of values that let a reader
    /// tell a right password from a wrong one without having the key to compare against.
    /// </summary>
    private static Locked Lock(Cipher key, string password, byte[] secretKey)
    {
        byte[] stretched = key.Hash.Spin(key.Hash.Of(key.Salt, Encoding.Unicode.GetBytes(password)), SpinCount);
        byte[] iv = Fit(key.Salt, key.BlockBytes);

        // The verifier is a random block; the file carries both it and its hash, each under a
        // key of its own. A reader that can decrypt the first and get the second has the
        // password right, and has learned nothing about the key the content is under.
        byte[] verifier = RandomNumberGenerator.GetBytes(SaltBytes);

        return new Locked(
            key.Encrypt(DeriveKey(key, stretched, VerifierInputBlock), iv, Pad(verifier, key.BlockBytes)),
            key.Encrypt(DeriveKey(key, stretched, VerifierValueBlock), iv, Pad(key.Hash.Of(verifier), key.BlockBytes)),
            key.Encrypt(DeriveKey(key, stretched, KeyValueBlock), iv, Pad(secretKey, key.BlockBytes)));
    }

    /// <summary>
    /// Produces the integrity pair: a random HMAC key, and the HMAC it gives over the finished
    /// package. Both travel encrypted under the secret key, so only a reader that already has
    /// the password can check them.
    /// </summary>
    private static (byte[] Key, byte[] Value) SealIntegrity(Cipher keyData, byte[] secretKey, byte[] encryptedPackage)
    {
        byte[] hmacKey = RandomNumberGenerator.GetBytes(keyData.Hash.Size);
        byte[] hmacValue = keyData.Hash.Hmac(hmacKey, encryptedPackage);

        return (
            keyData.Encrypt(secretKey, BlockIv(keyData, HmacKeyBlock), Pad(hmacKey, keyData.BlockBytes)),
            keyData.Encrypt(secretKey, BlockIv(keyData, HmacValueBlock), Pad(hmacValue, keyData.BlockBytes)));
    }

    /// <summary>
    /// Encrypts the package segment by segment, behind the length of the plaintext.
    /// </summary>
    /// <remarks>
    /// The last segment is written out to a whole cipher block rather than left short, and the
    /// length in front is what tells a reader where the content really ends. This is the exact
    /// inverse of <see cref="PackageBytes"/>: every segment but the last is 4096 bytes on both
    /// sides, so the segment number a reader computes from an offset is the one used here.
    /// </remarks>
    private static byte[] EncryptPackage(Cipher keyData, byte[] secretKey, ReadOnlySpan<byte> package)
    {
        int block = keyData.BlockBytes;
        int total = 0;
        for (int offset = 0; offset < package.Length; offset += SegmentSize)
            total += RoundUp(Math.Min(SegmentSize, package.Length - offset), block);

        var stream = new byte[8 + total];
        BinaryPrimitives.WriteInt64LittleEndian(stream, package.Length);

        Span<byte> segment = stackalloc byte[SegmentSize];
        for (int offset = 0, index = 0, written = 8; offset < package.Length; offset += SegmentSize, index++)
        {
            int length = Math.Min(SegmentSize, package.Length - offset);
            int padded = RoundUp(length, block);
            segment[..padded].Clear();
            package.Slice(offset, length).CopyTo(segment);

            byte[] iv = Fit(keyData.Hash.Of(keyData.Salt, index), block);
            keyData.Encrypt(secretKey, iv, segment[..padded]).CopyTo(stream.AsSpan(written));
            written += padded;
        }

        return stream;
    }

    /// <summary>
    /// Writes the descriptor: what the content is encrypted with, how to check it survived, and
    /// how a password gets at the key (MS-OFFCRYPTO 2.3.4.10).
    /// </summary>
    /// <remarks>
    /// The namespaces matter even though this library's own reader matches on local names
    /// alone: Excel and everything else validate against the schema, and the password encryptor
    /// is identified by its <c>uri</c>, not by where it sits.
    /// </remarks>
    private static byte[] Describe(Cipher keyData, Cipher passwordKey, Locked locked, byte[] hmacKey, byte[] hmacValue)
    {
        const string PasswordUri = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
        var xml = new StringBuilder(2048);

        xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n")
            .Append("<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\"")
            .Append(" xmlns:p=\"").Append(PasswordUri).Append('"')
            .Append(" xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">");

        xml.Append("<keyData");
        Parameters(xml, keyData);
        xml.Append("/>");

        xml.Append("<dataIntegrity encryptedHmacKey=\"").Append(Convert.ToBase64String(hmacKey))
            .Append("\" encryptedHmacValue=\"").Append(Convert.ToBase64String(hmacValue))
            .Append("\"/>");

        xml.Append("<keyEncryptors><keyEncryptor uri=\"").Append(PasswordUri).Append("\">")
            .Append("<p:encryptedKey spinCount=\"").Append(SpinCount.ToString(CultureInfo.InvariantCulture)).Append('"');
        Parameters(xml, passwordKey);
        xml.Append(" encryptedVerifierHashInput=\"").Append(Convert.ToBase64String(locked.VerifierInput))
            .Append("\" encryptedVerifierHashValue=\"").Append(Convert.ToBase64String(locked.VerifierValue))
            .Append("\" encryptedKeyValue=\"").Append(Convert.ToBase64String(locked.EncryptedKey))
            .Append("\"/></keyEncryptor></keyEncryptors></encryption>");

        return Encoding.UTF8.GetBytes(xml.ToString());
    }

    /// <summary>The algorithm attributes, which both the content and the password key declare.</summary>
    private static void Parameters(StringBuilder xml, Cipher cipher)
    {
        xml.Append(" saltSize=\"").Append(cipher.Salt.Length.ToString(CultureInfo.InvariantCulture))
            .Append("\" blockSize=\"").Append(cipher.BlockBytes.ToString(CultureInfo.InvariantCulture))
            .Append("\" keyBits=\"").Append((cipher.KeyBytes * 8).ToString(CultureInfo.InvariantCulture))
            .Append("\" hashSize=\"").Append(cipher.Hash.Size.ToString(CultureInfo.InvariantCulture))
            .Append("\" cipherAlgorithm=\"").Append(Name(cipher.Kind))
            .Append("\" cipherChaining=\"").Append(cipher.Cfb ? "ChainingModeCFB" : "ChainingModeCBC")
            .Append("\" hashAlgorithm=\"").Append(cipher.Hash.Algorithm.Name)
            .Append("\" saltValue=\"").Append(Convert.ToBase64String(cipher.Salt))
            .Append('"');
    }

    /// <summary>Zero-pads a value out to a whole cipher block, which is all these ciphers take.</summary>
    private static byte[] Pad(byte[] data, int blockBytes)
    {
        int padded = RoundUp(data.Length, blockBytes);
        if (padded == data.Length)
            return data;

        var buffer = new byte[padded];
        data.CopyTo(buffer, 0);
        return buffer;
    }

    private static int RoundUp(int value, int multiple) => (value + multiple - 1) / multiple * multiple;

    /// <summary>
    /// Derives the key that actually encrypts the package, or <see langword="null"/> when the
    /// password is not the one. The password never encrypts the content itself: it unlocks a
    /// random key, which is what lets one file be opened by several credentials.
    /// </summary>
    private static byte[]? TryUnlock(
        Descriptor descriptor,
        string password,
        CancellationToken cancellationToken)
    {
        Cipher key = descriptor.PasswordKey;
        byte[] stretched = key.Hash.Spin(
            key.Hash.Of(key.Salt, Encoding.Unicode.GetBytes(password)),
            descriptor.SpinCount,
            cancellationToken);

        // Every operation of the password key encryptor takes the salt itself as its
        // initialization vector, cut or padded to the cipher's block (2.3.4.12 case 2).
        byte[] iv = Fit(key.Salt, key.BlockBytes);

        try
        {
            byte[] verifier = key.Decrypt(DeriveKey(key, stretched, VerifierInputBlock), iv, descriptor.VerifierInput);
            byte[] expected = key.Hash.Of(verifier.AsSpan(0, Math.Min(verifier.Length, key.Salt.Length)));
            byte[] actual = key.Decrypt(DeriveKey(key, stretched, VerifierValueBlock), iv, descriptor.VerifierValue);
            if (actual.Length < expected.Length ||
                !CryptographicOperations.FixedTimeEquals(actual.AsSpan(0, expected.Length), expected))
            {
                return null;
            }

            byte[] unwrapped = key.Decrypt(DeriveKey(key, stretched, KeyValueBlock), iv, descriptor.EncryptedKey);
            return unwrapped.Length >= descriptor.KeyData.KeyBytes ? unwrapped[..descriptor.KeyData.KeyBytes] : null;
        }
        catch (CryptographicException)
        {
            // A wrong password produces a random key, and DES and its relatives reject a key
            // that happens to be one of their weak ones. That is a failed attempt, not a fault.
            return null;
        }
    }

    /// <summary>The key for one purpose: the stretched password hashed with that purpose's block key.</summary>
    private static byte[] DeriveKey(Cipher key, byte[] stretched, ReadOnlySpan<byte> blockKey) =>
        Fit(key.Hash.Of(stretched, blockKey), key.KeyBytes);

    /// <summary>
    /// Checks the package against the HMAC the file carries (MS-OFFCRYPTO 2.3.4.14). A file
    /// that fails this has been altered since it was written; the content is still handed back,
    /// because refusing it would help nobody, but the caller is told.
    /// </summary>
    /// <remarks>
    /// The key is stored encrypted and padded out to a whole cipher block, so its own length is
    /// not recoverable from the file. The specification says it is as long as the salt; every
    /// file examined here keys the HMAC with a hash-sized value instead. Since the answer is a
    /// warning rather than a refusal, the plausible lengths are all tried, and only a file that
    /// matches under none of them is reported — a false accusation of tampering would be worse
    /// than a miss.
    /// </remarks>
    private static void VerifyIntegrity(
        Descriptor descriptor, byte[] secretKey, byte[] encryptedPackage, Action<WorkbookWarning>? warn)
    {
        if (warn is null || descriptor.HmacKey is not { } encryptedHmacKey || descriptor.HmacValue is not { } encryptedHmacValue)
            return;

        Cipher data = descriptor.KeyData;
        byte[] hmacKey, expected;
        try
        {
            hmacKey = data.Decrypt(secretKey, BlockIv(data, HmacKeyBlock), encryptedHmacKey);
            expected = data.Decrypt(secretKey, BlockIv(data, HmacValueBlock), encryptedHmacValue);
        }
        catch (CryptographicException)
        {
            return;
        }

        int hashSize = Math.Min(data.Hash.Size, expected.Length);
        foreach (int length in new[] { data.Hash.Size, data.Salt.Length, hmacKey.Length }.Distinct())
        {
            if (length <= 0 || length > hmacKey.Length)
                continue;

            byte[] actual = data.Hash.Hmac(hmacKey[..length], encryptedPackage);
            if (CryptographicOperations.FixedTimeEquals(actual.AsSpan(0, hashSize), expected.AsSpan(0, hashSize)))
                return;
        }

        warn(new WorkbookWarning(
            "PackageIntegrityFailed",
            "The encrypted package does not match the integrity value stored with it, so it has been " +
            "altered since it was written. It was decrypted anyway."));
    }

    /// <summary>An initialization vector for one purpose (MS-OFFCRYPTO 2.3.4.12 case 1).</summary>
    private static byte[] BlockIv(Cipher cipher, ReadOnlySpan<byte> blockKey) =>
        Fit(cipher.Hash.Of(cipher.Salt, blockKey), cipher.BlockBytes);

    /// <summary>Decrypts the package segment by segment, then cuts it to the length it declares.</summary>
    private static byte[] PackageBytes(
        Cipher keyData,
        byte[] secretKey,
        byte[] encryptedPackage,
        CancellationToken cancellationToken)
    {
        if (encryptedPackage.Length < 8)
            throw new XlsFormatException("The encrypted package stream is too short to hold its own length.");

        long declared = BinaryPrimitives.ReadInt64LittleEndian(encryptedPackage);
        int available = encryptedPackage.Length - 8;
        if (declared is < 0 || declared > available)
            declared = available;

        int block = keyData.BlockBytes;
        var plain = new byte[available];
        Span<byte> segment = stackalloc byte[SegmentSize + Cipher.MaxBlockBytes];
        for (int offset = 0, index = 0; offset < available; offset += SegmentSize, index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(SegmentSize, available - offset);

            // A trailing segment shorter than the cipher block is padded rather than skipped:
            // those bytes are still ciphertext, and the declared length decides what of it counts.
            int padded = (length + block - 1) / block * block;
            segment[..padded].Clear();
            encryptedPackage.AsSpan(8 + offset, length).CopyTo(segment);

            byte[] iv = Fit(keyData.Hash.Of(keyData.Salt, index), block);
            keyData.Decrypt(secretKey, iv, segment[..padded]).AsSpan(0, length).CopyTo(plain.AsSpan(offset));
        }

        return declared == available ? plain : plain[..(int)declared];
    }

    /// <summary>
    /// Truncates a digest to a key or vector length, padding a short one as the spec says. The
    /// same rule governs both (2.3.4.11 and 2.3.4.12), and the padding byte is not zero.
    /// </summary>
    private static byte[] Fit(ReadOnlySpan<byte> digest, int length)
    {
        if (digest.Length == length)
            return digest.ToArray();
        if (digest.Length > length)
            return digest[..length].ToArray();

        var padded = new byte[length];
        padded.AsSpan().Fill(KeyPadding);
        digest.CopyTo(padded);
        return padded;
    }

    /// <summary>Everything the descriptor says, reduced to what decryption needs.</summary>
    private readonly record struct Descriptor(
        Cipher KeyData,
        Cipher PasswordKey,
        int SpinCount,
        byte[] VerifierInput,
        byte[] VerifierValue,
        byte[] EncryptedKey,
        byte[]? HmacKey,
        byte[]? HmacValue);

    /// <summary>The block ciphers a descriptor may name (MS-OFFCRYPTO 2.3.4.10).</summary>
    private enum CipherKind
    {
        Aes,
        Des,
        TripleDes,
        Rc2,
    }

    /// <summary>One set of algorithm parameters; the file declares them twice, for the two keys.</summary>
    private readonly record struct Cipher(
        CryptoHash Hash, CipherKind Kind, int KeyBytes, int BlockBytes, bool Cfb, byte[] Salt)
    {
        /// <summary>The widest block any of these ciphers has, for stack buffers.</summary>
        public const int MaxBlockBytes = 16;

        public byte[] Decrypt(byte[] key, ReadOnlySpan<byte> iv, ReadOnlySpan<byte> data)
        {
            using SymmetricAlgorithm algorithm = Create(key);

            // The specification fixes the feedback window at eight bits, which is not the
            // default any of these algorithms would pick.
            return Cfb
                ? algorithm.DecryptCfb(data, iv, PaddingMode.None, feedbackSizeInBits: 8)
                : algorithm.DecryptCbc(data, iv, PaddingMode.None);
        }

        /// <summary>
        /// The inverse. Padding is none here too: every caller has already brought its data up
        /// to a whole block, because the format decides what the extra bytes are rather than
        /// leaving it to a padding scheme a reader would have to strip.
        /// </summary>
        public byte[] Encrypt(byte[] key, ReadOnlySpan<byte> iv, ReadOnlySpan<byte> data)
        {
            using SymmetricAlgorithm algorithm = Create(key);
            return Cfb
                ? algorithm.EncryptCfb(data, iv, PaddingMode.None, feedbackSizeInBits: 8)
                : algorithm.EncryptCbc(data, iv, PaddingMode.None);
        }

        private SymmetricAlgorithm Create(byte[] key)
        {
            SymmetricAlgorithm algorithm = Kind switch
            {
                CipherKind.Des => DES.Create(),
                CipherKind.TripleDes => TripleDES.Create(),
                CipherKind.Rc2 => RC2.Create(),
                _ => Aes.Create(),
            };

            try
            {
                // RC2 is the one whose strength is separate from its key length, and Office
                // uses the whole key rather than weakening it.
                if (algorithm is RC2 rc2)
                    rc2.EffectiveKeySize = key.Length * 8;

                algorithm.Key = key;
                return algorithm;
            }
            catch
            {
                algorithm.Dispose();
                throw;
            }
        }
    }

    private static Descriptor Parse(
        ReadOnlySpan<byte> descriptor,
        DocumentLoadBudgetState? loadBudget,
        CancellationToken cancellationToken)
    {
        if (descriptor.Length > MaxDescriptorBytes)
        {
            throw new XlsFormatException(
                $"The agile encryption descriptor has {descriptor.Length} bytes; " +
                $"EncryptionInfo metadata is limited to {MaxDescriptorBytes}.");
        }

        Cipher? keyData = null;
        Cipher? passwordKey = null;
        int spinCount = 0;
        byte[]? verifierInput = null;
        byte[]? verifierValue = null;
        byte[]? encryptedKey = null;
        byte[]? hmacKey = null;
        byte[]? hmacValue = null;

        using var stream = new MemoryStream(descriptor.ToArray(), writable: false);
        XmlReaderSettings settings = loadBudget is null
            ? XmlDefaults.ReaderSettings
            : XmlDefaults.ForBudget(loadBudget.Budget, async: false);
        using XmlReader xml = XmlReader.Create(stream, settings);
        try
        {
            while (xml.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                loadBudget?.AddXmlNode(xml.Depth + 1);
                if (xml.NodeType != XmlNodeType.Element)
                    continue;

                switch (xml.LocalName)
                {
                    case "keyData":
                        keyData = ReadCipher(xml);
                        break;

                    case "dataIntegrity":
                        hmacKey = ReadBase64(xml, "encryptedHmacKey");
                        hmacValue = ReadBase64(xml, "encryptedHmacValue");
                        break;

                    // The password encryptor is one of possibly several; the others unlock the same
                    // file with a certificate, which is not a credential this library can supply.
                    case "encryptedKey":
                        passwordKey = ReadCipher(xml);
                        spinCount = ReadInt(xml, "spinCount");
                        verifierInput = ReadBase64(xml, "encryptedVerifierHashInput");
                        verifierValue = ReadBase64(xml, "encryptedVerifierHashValue");
                        encryptedKey = ReadBase64(xml, "encryptedKeyValue");
                        break;
                }
            }
        }
        catch (XmlException exception) when (
            loadBudget is not null &&
            exception.Message.Contains(
                nameof(XmlReaderSettings.MaxCharactersInDocument),
                StringComparison.Ordinal))
        {
            long limit = loadBudget.Budget.MaxXmlCharactersPerPart;
            throw new DocumentLoadLimitException(
                nameof(DocumentLoadBudget.MaxXmlCharactersPerPart),
                limit,
                limit == long.MaxValue ? long.MaxValue : limit + 1);
        }

        if (keyData is not { } data || passwordKey is not { } password)
        {
            throw new XlsFormatException(
                "The encryption descriptor names no password key, so this file cannot be opened with one.");
        }

        if (verifierInput is null || verifierValue is null || encryptedKey is null)
            throw new XlsFormatException("The encryption descriptor is missing the values that verify a password.");
        if (spinCount is < 0 or > 10_000_000)
            throw new XlsFormatException($"The encryption descriptor asks for {spinCount} hash rounds, which no real file does.");

        return new Descriptor(data, password, spinCount, verifierInput, verifierValue, encryptedKey, hmacKey, hmacValue);
    }

    private static Cipher ReadCipher(XmlReader xml)
    {
        string? hashName = xml.GetAttribute("hashAlgorithm");
        if (!CryptoHash.TryParse(hashName, out CryptoHash hash))
            throw new XlsFormatException($"The file is encrypted with the unsupported hash '{hashName}'.");

        string? algorithm = xml.GetAttribute("cipherAlgorithm");
        CipherKind kind = algorithm?.ToUpperInvariant() switch
        {
            "AES" => CipherKind.Aes,
            "DES" => CipherKind.Des,
            "3DES" or "3DES_112" => CipherKind.TripleDes,
            "RC2" => CipherKind.Rc2,
            _ => throw new XlsFormatException(
                $"The file is encrypted with '{algorithm}', which is not one of the block ciphers this library implements."),
        };

        string? chaining = xml.GetAttribute("cipherChaining");
        bool cfb = string.Equals(chaining, "ChainingModeCFB", StringComparison.OrdinalIgnoreCase);
        if (!cfb && !string.Equals(chaining, "ChainingModeCBC", StringComparison.OrdinalIgnoreCase))
            throw new XlsFormatException($"The file is encrypted in '{chaining}' mode, which is neither CBC nor CFB.");

        int keyBytes = KeyBytes(kind, ReadInt(xml, "keyBits"));
        int blockBytes = BlockBytes(kind, ReadInt(xml, "blockSize"));

        byte[] salt = ReadBase64(xml, "saltValue") ??
            throw new XlsFormatException("The encryption descriptor carries no salt.");

        return new Cipher(hash, kind, keyBytes, blockBytes, cfb, salt);
    }

    /// <summary>
    /// Key length in bytes, checked against what the named cipher actually accepts. DES counts
    /// its parity bits, so its 56-bit key is declared as 64 (MS-OFFCRYPTO 2.3.4.10, note 16).
    /// </summary>
    private static int KeyBytes(CipherKind kind, int keyBits)
    {
        bool valid = kind switch
        {
            CipherKind.Aes => keyBits is 128 or 192 or 256,
            CipherKind.Des => keyBits is 64,
            CipherKind.TripleDes => keyBits is 128 or 192,
            _ => keyBits is >= 40 and <= 128 && keyBits % 8 == 0,
        };

        return valid
            ? keyBits / 8
            : throw new XlsFormatException($"The file declares a {keyBits}-bit {Name(kind)} key, which is not a size that cipher has.");
    }

    /// <summary>
    /// Block size in bytes. The descriptor states it, but it is a property of the cipher, and a
    /// file that disagrees with its own algorithm would decrypt to noise.
    /// </summary>
    private static int BlockBytes(CipherKind kind, int declared)
    {
        int actual = kind == CipherKind.Aes ? 16 : 8;
        return declared is 0 || declared == actual
            ? actual
            : throw new XlsFormatException($"The file declares a {declared}-byte block, and {Name(kind)} has {actual}.");
    }

    private static string Name(CipherKind kind) => kind switch
    {
        CipherKind.Des => "DES",
        CipherKind.TripleDes => "3DES",
        CipherKind.Rc2 => "RC2",
        _ => "AES",
    };

    private static int ReadInt(XmlReader xml, string attribute) =>
        int.TryParse(xml.GetAttribute(attribute), CultureInfo.InvariantCulture, out int value) ? value : 0;

    private static byte[]? ReadBase64(XmlReader xml, string attribute)
    {
        if (xml.GetAttribute(attribute) is not { Length: > 0 } text)
            return null;

        if (text.Length > MaxEncodedCryptoValueCharacters)
        {
            throw new XlsFormatException(
                $"The encryption descriptor's '{attribute}' value has {text.Length} encoded characters; " +
                $"cryptographic metadata is limited to {MaxEncodedCryptoValueCharacters}.");
        }

        var buffer = new byte[(text.Length / 4 * 3) + 3];
        return Convert.TryFromBase64String(text, buffer, out int written) ? buffer[..written] : null;
    }
}
