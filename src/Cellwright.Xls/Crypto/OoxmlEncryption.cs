using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using Cellwright.Xls.Cfb;

namespace Cellwright.Xls.Crypto;

/// <summary>
/// Undoes the "password to open" protection of a modern package. Encrypting an .xlsx does not
/// produce an encrypted zip: Excel puts the whole package inside a compound file next to a
/// stream describing how it was encrypted — which is why such a file looks like a legacy .xls
/// to anything that reads only the first eight bytes.
/// </summary>
internal static class OoxmlEncryption
{
    /// <summary>The password Excel applies when a file is encrypted but not password-protected.</summary>
    public const string DefaultPassword = "VelvetSweatshop";

    private const string InfoStream = "EncryptionInfo";
    private const string PackageStream = "EncryptedPackage";

    /// <summary>Whether the container holds an encrypted package rather than a BIFF workbook.</summary>
    public static bool IsEncryptedPackage(CompoundFile container)
    {
        bool info = false;
        bool package = false;
        foreach (string name in container.EntryNames)
        {
            info |= string.Equals(name, InfoStream, StringComparison.Ordinal);
            package |= string.Equals(name, PackageStream, StringComparison.Ordinal);
        }

        return info && package;
    }

    /// <summary>Recovers the package inside the container.</summary>
    /// <param name="container">A container <see cref="IsEncryptedPackage"/> accepted.</param>
    /// <param name="password">Password to try; the default one is used when none is given.</param>
    /// <param name="warn">Reports a package that fails the integrity check stored with it.</param>
    /// <param name="loadBudget">Shared counters inherited by the decrypted package.</param>
    /// <param name="cancellationToken">Cancels password stretching and package decryption.</param>
    /// <exception cref="XlsFormatException">The scheme is unsupported, or no password opens the file.</exception>
    public static byte[] Decrypt(
        CompoundFile container,
        string? password,
        Action<WorkbookWarning>? warn = null,
        DocumentLoadBudgetState? loadBudget = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        long? infoLength = container.TryGetStreamLength(InfoStream);
        if (infoLength is null)
            throw new XlsFormatException("The encrypted file has no encryption header.");
        long maxInfoBytes = AgileEncryption.MaxDescriptorBytes + 8L;
        if (infoLength > maxInfoBytes)
        {
            throw new XlsFormatException(
                $"The EncryptionInfo stream has {infoLength} bytes; encryption metadata is limited to {maxInfoBytes}.");
        }

        byte[] info = container.TryReadStream(InfoStream) ??
            throw new XlsFormatException("The encrypted file has no encryption header.");
        byte[] package = container.TryReadStream(PackageStream) ??
            throw new XlsFormatException("The encrypted file has no package stream.");
        if (info.Length < 8)
            throw new XlsFormatException("The encryption header is too short to say how the file is encrypted.");

        ushort major = BinaryPrimitives.ReadUInt16LittleEndian(info);
        ushort minor = BinaryPrimitives.ReadUInt16LittleEndian(info.AsSpan(2));
        long nestedInput = Math.Max(0, package.LongLength - 8);
        if (package.Length >= 8)
        {
            long declared = BinaryPrimitives.ReadInt64LittleEndian(package);
            if (declared >= 0)
                nestedInput = Math.Max(nestedInput, declared);
        }
        if (loadBudget is not null)
        {
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxInputBytes),
                loadBudget.Budget.MaxInputBytes,
                nestedInput);
        }

        // 4.4 is agile encryption, whose header is an XML descriptor; 2.2 through 4.2 are the
        // fixed AES scheme of Office 2007. Minor 3 is the extensible form, which delegates the
        // whole thing to a third-party module and so cannot be read without it.
        byte[]? plain = (major, minor) switch
        {
            (4, 4) => AgileEncryption.Decrypt(
                info.AsSpan(8), package, Candidates(password), warn, cancellationToken, loadBudget),
            (2 or 3 or 4, 2) => DecryptStandard(
                info.AsSpan(8), package, Candidates(password), cancellationToken),
            _ => throw new XlsFormatException(
                $"The file uses encryption version {major}.{minor}, which this library does not implement."),
        };

        return plain ?? throw new XlsFormatException(PasswordFailure(password));
    }

    /// <summary>
    /// Puts a package behind a password, as the two streams a container holds it in.
    /// </summary>
    /// <remarks>
    /// The inverse of <see cref="Decrypt"/>, and deliberately narrower: agile encryption is the
    /// only scheme written. The standard one is readable because files exist that use it, not
    /// because anything should still be producing it.
    /// </remarks>
    /// <param name="package">The plain package.</param>
    /// <param name="password">The password that will open it.</param>
    public static IReadOnlyList<CompoundFileWriter.NamedStream> Encrypt(ReadOnlySpan<byte> package, string password)
    {
        AgileEncryption.Written written = AgileEncryption.Encrypt(package, password);

        // The version header in front of the descriptor: 4.4 is agile, and the flags say the
        // rest is an XML descriptor rather than the fixed structures of the older schemes.
        var info = new byte[8 + written.Descriptor.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(info, 4);
        BinaryPrimitives.WriteUInt16LittleEndian(info.AsSpan(2), 4);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(4), 0x40);
        written.Descriptor.CopyTo(info.AsSpan(8));

        return
        [
            new CompoundFileWriter.NamedStream(InfoStream, info),
            new CompoundFileWriter.NamedStream(PackageStream, written.EncryptedPackage),
        ];
    }

    /// <summary>What to say when nothing opened the file, which depends on whether anything was tried.</summary>
    public static string PasswordFailure(string? password) => string.IsNullOrEmpty(password)
        ? "The workbook is encrypted and needs a password; set XlsLoadOptions.Password."
        : "The password does not open this workbook.";

    /// <summary>The passwords to try, default first, so that a silently encrypted file just opens.</summary>
    private static IEnumerable<string> Candidates(string? password)
    {
        yield return DefaultPassword;
        if (!string.IsNullOrEmpty(password) && password != DefaultPassword)
            yield return password;
    }

    /// <summary>
    /// The scheme Office 2007 wrote: AES in ECB mode over the whole package, with a key derived
    /// from fifty thousand rounds of SHA-1. Everything about it is fixed except the key length,
    /// so the header is read only far enough to confirm that it is AES and how wide.
    /// </summary>
    private static byte[]? DecryptStandard(
        ReadOnlySpan<byte> info,
        byte[] encryptedPackage,
        IEnumerable<string> passwords,
        CancellationToken cancellationToken)
    {
        if (info.Length < 4)
            throw new XlsFormatException("The encryption header declares no size.");

        // The verifier that follows the header is 72 bytes: a salt, sixteen encrypted random
        // bytes, and their SHA-1 padded out to two cipher blocks.
        int headerSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(info);
        if (headerSize < 32 || 4 + headerSize + 72 > info.Length)
            throw new XlsFormatException("The encryption header does not fit the stream that carries it.");

        ReadOnlySpan<byte> header = info.Slice(4, headerSize);
        uint algorithm = BinaryPrimitives.ReadUInt32LittleEndian(header[8..]);
        if (algorithm is not (0x660E or 0x660F or 0x6610))
            throw new XlsFormatException($"The package is encrypted with algorithm 0x{algorithm:X4}, and only AES is supported.");

        int keyBytes = (int)BinaryPrimitives.ReadUInt32LittleEndian(header[16..]) / 8;
        if (keyBytes is not (16 or 24 or 32))
            throw new XlsFormatException($"The package declares a {keyBytes * 8}-bit AES key, which is not a size AES has.");

        // EncryptionVerifier (MS-OFFCRYPTO 2.3.3): a salt, sixteen random bytes, and their hash,
        // the last two encrypted with the very key the password is supposed to produce.
        ReadOnlySpan<byte> verifier = info[(4 + headerSize)..];
        ReadOnlySpan<byte> salt = verifier.Slice(4, 16);
        ReadOnlySpan<byte> encryptedVerifier = verifier.Slice(20, 16);
        ReadOnlySpan<byte> encryptedHash = verifier.Slice(40, 32);

        foreach (string candidate in passwords)
        {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] key = StandardKey(candidate, salt, keyBytes, cancellationToken);
            if (VerifiesStandard(key, encryptedVerifier, encryptedHash))
                return StandardPackage(key, encryptedPackage, cancellationToken);
        }

        return null;
    }

    /// <summary>
    /// The key derivation of ECMA-376 standard encryption (MS-OFFCRYPTO 2.3.4.7): the password
    /// stretched by fifty thousand rounds, then run through the two-pad construction that turns
    /// one digest into as many key bytes as AES asks for.
    /// </summary>
    private static byte[] StandardKey(
        string password,
        ReadOnlySpan<byte> salt,
        int keyBytes,
        CancellationToken cancellationToken)
    {
        const int Rounds = 50_000;
        CryptoHash sha1 = CryptoHash.Sha1;
        byte[] stretched = sha1.Spin(
            sha1.Of(salt, Encoding.Unicode.GetBytes(password)), Rounds, cancellationToken);
        byte[] seed = sha1.Of(stretched, block: 0);

        Span<byte> first = stackalloc byte[64];
        Span<byte> second = stackalloc byte[64];
        first.Fill(0x36);
        second.Fill(0x5C);
        for (int i = 0; i < seed.Length; i++)
        {
            first[i] ^= seed[i];
            second[i] ^= seed[i];
        }

        Span<byte> derived = stackalloc byte[2 * 20];
        CryptographicOperations.HashData(sha1.Algorithm, first, derived[..20]);
        CryptographicOperations.HashData(sha1.Algorithm, second, derived[20..]);
        return derived[..keyBytes].ToArray();
    }

    private static bool VerifiesStandard(byte[] key, ReadOnlySpan<byte> encryptedVerifier, ReadOnlySpan<byte> encryptedHash)
    {
        using Aes aes = Aes.Create();
        aes.Key = key;
        byte[] verifier = aes.DecryptEcb(encryptedVerifier, PaddingMode.None);
        byte[] hash = aes.DecryptEcb(encryptedHash, PaddingMode.None);
        return CryptographicOperations.FixedTimeEquals(hash.AsSpan(0, CryptoHash.Sha1.Size), CryptoHash.Sha1.Of(verifier));
    }

    /// <summary>Decrypts the package stream, whose first eight bytes are the plaintext length.</summary>
    private static byte[] StandardPackage(
        byte[] key,
        byte[] encryptedPackage,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (encryptedPackage.Length < 8)
            throw new XlsFormatException("The encrypted package stream is too short to hold its own length.");

        long declared = BinaryPrimitives.ReadInt64LittleEndian(encryptedPackage);
        int available = (encryptedPackage.Length - 8) / 16 * 16;

        using Aes aes = Aes.Create();
        aes.Key = key;
        byte[] plain = aes.DecryptEcb(encryptedPackage.AsSpan(8, available), PaddingMode.None);
        cancellationToken.ThrowIfCancellationRequested();
        return declared is >= 0 && declared < plain.Length ? plain[..(int)declared] : plain;
    }
}
