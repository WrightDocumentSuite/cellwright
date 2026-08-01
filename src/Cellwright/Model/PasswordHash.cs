using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Cellwright;

/// <summary>
/// The verifier stored beside a protected sheet or workbook. Excel never keeps the password
/// itself: it keeps a hash and checks what the user types against it, and nothing about the
/// protection is encryption — anyone reading the file reads the data.
/// </summary>
/// <remarks>
/// Two forms exist. Old files hold a 16-bit checksum that a laptop brute-forces in a second;
/// they are read so that such a file survives a round trip. Anything written here uses the
/// iterated SHA-512 form Excel 2013 and later produce.
/// </remarks>
public sealed record PasswordHash
{
    private const int DefaultSpinCount = 100_000;
    private const string DefaultAlgorithm = "SHA-512";

    /// <summary>Base64 hash, or the four hex digits of the legacy checksum.</summary>
    public required string Value { get; init; }

    /// <summary>Base64 salt of the modern form; <see langword="null"/> for the legacy checksum.</summary>
    public string? Salt { get; init; }

    /// <summary>Name of the hash function, e.g. <c>SHA-512</c>; <see langword="null"/> for the legacy checksum.</summary>
    public string? Algorithm { get; init; }

    /// <summary>How many times the hash is re-applied; zero for the legacy checksum.</summary>
    public int SpinCount { get; init; }

    /// <summary>Whether this is the pre-2013 16-bit checksum rather than an iterated hash.</summary>
    public bool IsLegacy => Salt is null;

    /// <summary>Hashes a password the way Excel 2013 and later do.</summary>
    public static PasswordHash Create(string password)
    {
        ArgumentNullException.ThrowIfNull(password);
        byte[] salt = RandomNumberGenerator.GetBytes(16);
        return new PasswordHash
        {
            Value = Convert.ToBase64String(Spin(password, salt, DefaultSpinCount)),
            Salt = Convert.ToBase64String(salt),
            Algorithm = DefaultAlgorithm,
            SpinCount = DefaultSpinCount,
        };
    }

    /// <summary>Checks a password against the verifier, in either form.</summary>
    public bool Matches(string password)
    {
        ArgumentNullException.ThrowIfNull(password);
        // Producers disagree on padding and case, so the checksum is compared as a number.
        if (IsLegacy)
        {
            return ushort.TryParse(Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort stored) &&
                   stored == LegacyChecksum(password);
        }

        if (!string.Equals(Algorithm, DefaultAlgorithm, StringComparison.OrdinalIgnoreCase))
            return false;

        byte[] salt = Convert.FromBase64String(Salt!);
        return CryptographicOperations.FixedTimeEquals(
            Spin(password, salt, SpinCount), Convert.FromBase64String(Value));
    }

    /// <summary>Rebuilds a verifier read from a file.</summary>
    internal static PasswordHash? FromFile(string? hash, string? salt, string? algorithm, int spinCount, string? legacy)
    {
        if (!string.IsNullOrEmpty(hash))
            return new PasswordHash { Value = hash, Salt = salt, Algorithm = algorithm, SpinCount = spinCount };
        return string.IsNullOrEmpty(legacy) ? null : new PasswordHash { Value = legacy };
    }

    /// <summary>
    /// The iterated hash: SHA-512 of salt and password, then of the previous digest and the
    /// little-endian iteration counter, <paramref name="spinCount"/> times over.
    /// </summary>
    private static byte[] Spin(string password, byte[] salt, int spinCount)
    {
        byte[] encoded = Encoding.Unicode.GetBytes(password);
        byte[] digest = SHA512.HashData([.. salt, .. encoded]);

        Span<byte> buffer = stackalloc byte[64 + sizeof(int)];
        for (int iteration = 0; iteration < spinCount; iteration++)
        {
            digest.CopyTo(buffer);
            BitConverter.TryWriteBytes(buffer[64..], iteration);
            digest = SHA512.HashData(buffer);
        }

        return digest;
    }

    /// <summary>
    /// The 16-bit checksum of the pre-2013 format: rotate left through 15 bits and exclusive-or
    /// each character from the last, then fold in the length and the constant <c>0xCE4B</c>.
    /// </summary>
    private static ushort LegacyChecksum(string password)
    {
        if (password.Length == 0)
            return 0;

        int hash = 0;
        for (int i = password.Length - 1; i >= 0; i--)
        {
            hash = Rotate(hash);
            hash ^= password[i] & 0xFF;
        }

        hash = Rotate(hash);
        hash ^= password.Length;
        hash ^= 0x8000 | ('N' << 8) | 'K';
        return (ushort)hash;

        static int Rotate(int hash) => ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
    }
}
