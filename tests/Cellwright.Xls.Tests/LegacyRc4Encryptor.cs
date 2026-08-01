using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Encrypts a workbook stream the way Excel 97 did — RC4 with a key derived through MD5, the
/// scheme MS-OFFCRYPTO calls "RC4 encryption". No corpus file uses it and no library within
/// reach writes it, so the fixture is built here, from the specification rather than from the
/// reader it is meant to test.
/// </summary>
internal static class LegacyRc4Encryptor
{
    /// <summary>Bytes of the stream one key covers before the next takes over.</summary>
    private const int BlockSize = 1024;

    /// <summary>Records that stay in the clear, since a reader needs them before it has a key.</summary>
    private static bool StaysPlain(ushort recordType) => recordType is
        0x0809 or   // BOF.
        0x002F or   // FilePass.
        0x00E1 or   // InterfaceHdr.
        0x002A or   // WriteProt.
        0x0195;     // FileSharing.

    /// <summary>
    /// Rewrites a workbook stream as an encrypted one: the opening record is kept, a FilePass
    /// record is put after it, and everything from there on is enciphered.
    /// </summary>
    /// <param name="plain">The unencrypted workbook stream.</param>
    /// <param name="password">Password the file will need.</param>
    public static byte[] Encrypt(byte[] plain, string password)
    {
        byte[] salt = [.. Enumerable.Range(0, 16).Select(static i => (byte)(i * 7 + 3))];
        byte[] verifier = [.. Enumerable.Range(0, 16).Select(static i => (byte)(i * 11 + 5))];
        byte[] keyDigest = KeyDigest(password, salt);

        byte[] verifierHash = MD5.HashData(verifier);
        byte[] encryptedVerifier = [.. verifier];
        byte[] encryptedHash = [.. verifierHash];
        var header = new Rc4Cipher(BlockKey(keyDigest, 0));
        header.Apply(encryptedVerifier);
        header.Apply(encryptedHash);

        var output = new List<byte>(plain.Length + 64);
        int cursor = 0;

        // The opening record travels unchanged, and the encryption header follows it.
        int openingSize = BinaryPrimitives.ReadUInt16LittleEndian(plain.AsSpan(2));
        output.AddRange(plain.AsSpan(0, 4 + openingSize).ToArray());
        cursor += 4 + openingSize;
        output.AddRange(FilePassRecord(salt, encryptedVerifier, encryptedHash));

        while (cursor + 4 <= plain.Length)
        {
            ushort type = BinaryPrimitives.ReadUInt16LittleEndian(plain.AsSpan(cursor));
            int size = BinaryPrimitives.ReadUInt16LittleEndian(plain.AsSpan(cursor + 2));
            if (cursor + 4 + size > plain.Length)
                break;

            byte[] body = plain.AsSpan(cursor + 4, size).ToArray();
            int bodyOffset = output.Count + 4;
            if (!StaysPlain(type))
                EncryptBlocks(body, bodyOffset, keyDigest);

            // A reader has to find the sheets before it can decrypt anything, so the offset
            // that opens a BoundSheet8 record is left readable.
            if (type == 0x0085 && size >= 4)
                plain.AsSpan(cursor + 4, 4).CopyTo(body);

            output.AddRange(plain.AsSpan(cursor, 4).ToArray());
            output.AddRange(body);
            cursor += 4 + size;
        }

        return [.. output];
    }

    private static byte[] FilePassRecord(byte[] salt, byte[] verifier, byte[] verifierHash)
    {
        var body = new List<byte>(54);
        body.AddRange([0x01, 0x00]);        // RC4 rather than the XOR obfuscation.
        body.AddRange([0x01, 0x00]);        // Version 1.1: the original RC4 scheme.
        body.AddRange([0x01, 0x00]);
        body.AddRange(salt);
        body.AddRange(verifier);
        body.AddRange(verifierHash);

        var record = new List<byte>(58);
        record.AddRange([0x2F, 0x00]);
        record.AddRange([(byte)body.Count, 0x00]);
        record.AddRange(body);
        return [.. record];
    }

    /// <summary>Enciphers a record's body, which may straddle the boundary between two keys.</summary>
    private static void EncryptBlocks(byte[] body, int offset, byte[] keyDigest)
    {
        int written = 0;
        while (written < body.Length)
        {
            int position = offset + written;
            int block = position / BlockSize;
            int withinBlock = position % BlockSize;
            int take = Math.Min(BlockSize - withinBlock, body.Length - written);

            var cipher = new Rc4Cipher(BlockKey(keyDigest, block));
            cipher.Skip(withinBlock);
            cipher.Apply(body.AsSpan(written, take));
            written += take;
        }
    }

    /// <summary>
    /// The five bytes everything hangs off: the password hashed, then its first five bytes
    /// stirred with the salt sixteen times over (MS-OFFCRYPTO 2.3.6.2).
    /// </summary>
    private static byte[] KeyDigest(string password, byte[] salt)
    {
        byte[] hash = MD5.HashData(Encoding.Unicode.GetBytes(password));
        var material = new byte[16 * 21];
        for (int i = 0; i < 16; i++)
        {
            Array.Copy(hash, 0, material, i * 21, 5);
            Array.Copy(salt, 0, material, (i * 21) + 5, 16);
        }

        return [.. MD5.HashData(material).AsSpan(0, 5)];
    }

    /// <summary>The key of one block: the digest and the block number, hashed together.</summary>
    private static byte[] BlockKey(byte[] keyDigest, int block)
    {
        var material = new byte[keyDigest.Length + 4];
        keyDigest.CopyTo(material, 0);
        BinaryPrimitives.WriteInt32LittleEndian(material.AsSpan(keyDigest.Length), block);
        return MD5.HashData(material);
    }

    /// <summary>RC4, which the framework no longer offers and which this scheme still needs.</summary>
    private sealed class Rc4Cipher
    {
        private readonly byte[] _state = new byte[256];
        private int _i;
        private int _j;

        public Rc4Cipher(byte[] key)
        {
            for (int i = 0; i < 256; i++)
                _state[i] = (byte)i;

            int swap = 0;
            for (int i = 0; i < 256; i++)
            {
                swap = (swap + _state[i] + key[i % key.Length]) & 0xFF;
                (_state[i], _state[swap]) = (_state[swap], _state[i]);
            }
        }

        public void Skip(int count)
        {
            for (int n = 0; n < count; n++)
                NextByte();
        }

        public void Apply(Span<byte> data)
        {
            for (int n = 0; n < data.Length; n++)
                data[n] ^= NextByte();
        }

        private byte NextByte()
        {
            _i = (_i + 1) & 0xFF;
            _j = (_j + _state[_i]) & 0xFF;
            (_state[_i], _state[_j]) = (_state[_j], _state[_i]);
            return _state[(_state[_i] + _state[_j]) & 0xFF];
        }
    }
}
