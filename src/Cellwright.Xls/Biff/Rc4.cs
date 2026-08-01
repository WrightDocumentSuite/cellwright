namespace Cellwright.Xls.Biff;

/// <summary>
/// RC4, implemented here because the framework dropped it — rightly, since nothing should
/// encrypt with it today. Legacy workbooks were encrypted with it decades ago, and reading them
/// back is the only reason this exists.
/// </summary>
internal sealed class Rc4
{
    private readonly byte[] _state = new byte[256];
    private int _i;
    private int _j;

    /// <summary>Sets up the cipher from a key.</summary>
    public Rc4(ReadOnlySpan<byte> key)
    {
        for (int i = 0; i < 256; i++)
            _state[i] = (byte)i;

        int j = 0;
        for (int i = 0; i < 256; i++)
        {
            j = (j + _state[i] + key[i % key.Length]) & 0xFF;
            (_state[i], _state[j]) = (_state[j], _state[i]);
        }
    }

    /// <summary>Exclusive-ors the keystream over a buffer, in place.</summary>
    public void Apply(Span<byte> buffer)
    {
        for (int index = 0; index < buffer.Length; index++)
            buffer[index] ^= NextKeyByte();
    }

    /// <summary>Advances the keystream without touching any data, to skip over bytes.</summary>
    public void Skip(int count)
    {
        for (int index = 0; index < count; index++)
            NextKeyByte();
    }

    private byte NextKeyByte()
    {
        _i = (_i + 1) & 0xFF;
        _j = (_j + _state[_i]) & 0xFF;
        (_state[_i], _state[_j]) = (_state[_j], _state[_i]);
        return _state[(_state[_i] + _state[_j]) & 0xFF];
    }
}
