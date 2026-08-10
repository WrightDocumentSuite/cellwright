using System.Buffers.Binary;
using System.Security.Cryptography;

namespace Cellwright.Xls.Crypto;

/// <summary>
/// The hash function a document's encryption was built on. Agile encryption names it in the
/// file rather than fixing it, so every algorithm Office allows has to be reachable by name.
/// </summary>
/// <param name="Algorithm">Which hash to run.</param>
/// <param name="Size">Length of its digest in bytes; zero for an algorithm we do not know.</param>
internal readonly record struct CryptoHash(HashAlgorithmName Algorithm, int Size)
{
    /// <summary>The hash of the original ECMA-376 encryption, which allows nothing else.</summary>
    public static CryptoHash Sha1 => new(HashAlgorithmName.SHA1, 20);

    /// <summary>The hash written into every file this library encrypts.</summary>
    public static CryptoHash Sha512 => new(HashAlgorithmName.SHA512, 64);

    /// <summary>The longest digest any of these algorithms produces, for stack buffers.</summary>
    public const int MaxSize = 64;

    /// <summary>Resolves the name a file uses for its hash.</summary>
    public static bool TryParse(string? name, out CryptoHash hash)
    {
        hash = name?.ToUpperInvariant() switch
        {
            "SHA1" or "SHA-1" => Sha1,
            "SHA256" or "SHA-256" => new CryptoHash(HashAlgorithmName.SHA256, 32),
            "SHA384" or "SHA-384" => new CryptoHash(HashAlgorithmName.SHA384, 48),
            "SHA512" or "SHA-512" => new CryptoHash(HashAlgorithmName.SHA512, 64),
            "MD5" => new CryptoHash(HashAlgorithmName.MD5, 16),
            _ => default,
        };

        return hash.Size > 0;
    }

    /// <summary>Hashes one buffer.</summary>
    public byte[] Of(ReadOnlySpan<byte> data)
    {
        var digest = new byte[Size];
        CryptographicOperations.HashData(Algorithm, data, digest);
        return digest;
    }

    /// <summary>Hashes two buffers as if they were one.</summary>
    public byte[] Of(ReadOnlySpan<byte> first, ReadOnlySpan<byte> second)
    {
        Span<byte> material = new byte[first.Length + second.Length];
        first.CopyTo(material);
        second.CopyTo(material[first.Length..]);
        return Of(material);
    }

    /// <summary>Hashes a buffer with a little-endian block number appended.</summary>
    public byte[] Of(ReadOnlySpan<byte> data, int block)
    {
        Span<byte> number = stackalloc byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(number, block);
        return Of(data, number);
    }

    /// <summary>Keyed hash of a message, for the integrity value of an encrypted package.</summary>
    public byte[] Hmac(byte[] key, ReadOnlySpan<byte> message)
    {
        using IncrementalHash hmac = IncrementalHash.CreateHMAC(Algorithm, key);
        hmac.AppendData(message);
        return hmac.GetHashAndReset();
    }

    /// <summary>
    /// Runs the iteration both encryption schemes stretch a password with:
    /// <c>H(i) = Hash(LE32(i) || H(i-1))</c>, a hundred thousand times over in a modern file.
    /// One buffer serves the whole loop, since the digest goes straight back into the material.
    /// </summary>
    /// <param name="seed">The hash of the salt and the password.</param>
    /// <param name="count">How many rounds the file asks for.</param>
    /// <param name="cancellationToken">Cancels a long password-stretching loop.</param>
    public byte[] Spin(
        ReadOnlySpan<byte> seed,
        int count,
        CancellationToken cancellationToken = default)
    {
        Span<byte> material = stackalloc byte[4 + MaxSize];
        Span<byte> digest = stackalloc byte[MaxSize];
        seed[..Size].CopyTo(material[4..]);

        for (int i = 0; i < count; i++)
        {
            if ((i & 1023) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            BinaryPrimitives.WriteInt32LittleEndian(material, i);
            CryptographicOperations.HashData(Algorithm, material[..(4 + Size)], digest[..Size]);
            digest[..Size].CopyTo(material[4..]);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return material.Slice(4, Size).ToArray();
    }
}
