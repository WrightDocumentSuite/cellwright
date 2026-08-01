using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.Text.Unicode;

namespace Cellwright.Xml;

/// <summary>
/// Minimal forward-only XML writer that emits UTF-8 bytes directly into a pooled buffer.
/// Element structure is the caller's responsibility (via <see cref="WriteRaw"/> of u8 literals);
/// this type owns correct text escaping, number formatting and buffered async flushing.
/// </summary>
internal sealed class Utf8XmlWriter : IAsyncDisposable
{
    private const int DefaultBufferSize = 64 * 1024;
    private const int FlushThreshold = 48 * 1024;

    // '&', '<', '>', '"' plus every XML 1.0-invalid control character; '_' starts OOXML _xHHHH_ escapes.
    private static readonly SearchValues<char> NeedsEscape = SearchValues.Create(
        "&<>\"_\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C\u000E\u000F" +
        "\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F" +
        "\t\n\r");

    private readonly Stream _stream;
    private byte[] _buffer;
    private int _position;

    public Utf8XmlWriter(Stream stream, int bufferSize = DefaultBufferSize)
    {
        _stream = stream;
        _buffer = ArrayPool<byte>.Shared.Rent(bufferSize);
    }

    /// <summary>Writes the standard xlsx XML declaration.</summary>
    public void WriteDeclaration() =>
        WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);

    /// <summary>Writes pre-encoded UTF-8 markup verbatim.</summary>
    public void WriteRaw(scoped ReadOnlySpan<byte> utf8)
    {
        EnsureCapacity(utf8.Length);
        utf8.CopyTo(_buffer.AsSpan(_position));
        _position += utf8.Length;
    }

    /// <summary>
    /// Writes already-formed markup, transcoding UTF-16 to UTF-8 without escaping. Used to
    /// re-emit parts of a document that were captured verbatim while reading.
    /// </summary>
    public void WriteRawXml(scoped ReadOnlySpan<char> xml) => TranscodeChunk(xml);

    /// <summary>Writes escaped element text.</summary>
    public void WriteText(scoped ReadOnlySpan<char> text) => WriteEscaped(text, attribute: false);

    /// <summary>Writes an escaped attribute value (quotes and whitespace get character references).</summary>
    public void WriteAttributeText(scoped ReadOnlySpan<char> text) => WriteEscaped(text, attribute: true);

    /// <summary>Writes an integer in invariant culture.</summary>
    public void WriteInt32(int value)
    {
        EnsureCapacity(11);
        Utf8Formatter.TryFormat(value, _buffer.AsSpan(_position), out int written);
        _position += written;
    }

    /// <summary>Writes a long integer in invariant culture, e.g. a drawing measurement in EMU.</summary>
    public void WriteInt64(long value)
    {
        EnsureCapacity(20);
        Utf8Formatter.TryFormat(value, _buffer.AsSpan(_position), out int written);
        _position += written;
    }

    /// <summary>Writes a number using the shortest round-trippable invariant representation.</summary>
    public void WriteDouble(double value)
    {
        EnsureCapacity(32);
        value.TryFormat(_buffer.AsSpan(_position), out int written, default, CultureInfo.InvariantCulture);
        _position += written;
    }

    /// <summary>Writes any UTF-8 formattable value, e.g. a <see cref="CellRef"/>.</summary>
    public void WriteFormattable<T>(T value) where T : IUtf8SpanFormattable
    {
        EnsureCapacity(32);
        if (!value.TryFormat(_buffer.AsSpan(_position), out int written, default, CultureInfo.InvariantCulture))
        {
            EnsureCapacity(256);
            value.TryFormat(_buffer.AsSpan(_position), out written, default, CultureInfo.InvariantCulture);
        }

        _position += written;
    }

    /// <summary>Flushes the buffer to the underlying stream when it crossed the threshold.</summary>
    public ValueTask FlushIfNeededAsync(CancellationToken cancellationToken) =>
        _position >= FlushThreshold ? FlushAsync(cancellationToken) : ValueTask.CompletedTask;

    /// <summary>Flushes all buffered bytes to the underlying stream.</summary>
    public async ValueTask FlushAsync(CancellationToken cancellationToken)
    {
        if (_position > 0)
        {
            await _stream.WriteAsync(_buffer.AsMemory(0, _position), cancellationToken).ConfigureAwait(false);
            _position = 0;
        }
    }

    /// <summary>
    /// Renders a small fragment to a string, for the callers that need markup in hand rather than
    /// in a part — a cell's rich-text runs kept beside the model, say.
    /// </summary>
    /// <param name="write">What to write.</param>
    /// <remarks>
    /// Synchronous on purpose, and safe to be: nothing here touches a stream until it is asked to.
    /// Writing only fills a pooled buffer, which grows rather than flushing when it runs out, so
    /// the single write to the <see cref="MemoryStream"/> happens at the end under this method's
    /// own control. That is also why this is for fragments and not for parts: the whole thing is
    /// held in memory at once, which is the opposite of what the streaming path is for.
    /// </remarks>
    public static string Render(Action<Utf8XmlWriter> write)
    {
        using var stream = new MemoryStream();
        var writer = new Utf8XmlWriter(stream, bufferSize: 1024);
        try
        {
            write(writer);
            writer.FlushToStream();
        }
        finally
        {
            writer.ReturnBuffer();
        }

        return System.Text.Encoding.UTF8.GetString(stream.GetBuffer(), 0, (int)stream.Length);
    }

    /// <summary>Writes the buffered bytes straight through, without going async.</summary>
    private void FlushToStream()
    {
        if (_position == 0)
            return;

        _stream.Write(_buffer.AsSpan(0, _position));
        _position = 0;
    }

    /// <summary>Gives the pooled buffer back. The writer is unusable afterwards.</summary>
    private void ReturnBuffer()
    {
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = [];
    }

    /// <summary>Flushes remaining bytes, returns the pooled buffer and disposes the underlying stream.</summary>
    public async ValueTask DisposeAsync()
    {
        await FlushAsync(CancellationToken.None).ConfigureAwait(false);
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = [];
        await _stream.DisposeAsync().ConfigureAwait(false);
    }

    private void WriteEscaped(scoped ReadOnlySpan<char> text, bool attribute)
    {
        while (!text.IsEmpty)
        {
            int special = text.IndexOfAny(NeedsEscape);
            if (special < 0)
            {
                TranscodeChunk(text);
                return;
            }

            if (special > 0)
                TranscodeChunk(text[..special]);

            char c = text[special];
            int consumed = WriteSpecial(c, text[special..], attribute);
            text = text[(special + consumed)..];
        }
    }

    /// <summary>Writes one special character (or passes through a benign underscore). Returns chars consumed.</summary>
    private int WriteSpecial(char c, scoped ReadOnlySpan<char> rest, bool attribute)
    {
        switch (c)
        {
            case '&':
                WriteRaw("&amp;"u8);
                return 1;
            case '<':
                WriteRaw("&lt;"u8);
                return 1;
            case '>':
                WriteRaw("&gt;"u8);
                return 1;
            case '"':
                if (attribute) WriteRaw("&quot;"u8); else WriteByte((byte)'"');
                return 1;
            case '\t' or '\n' or '\r':
                if (attribute)
                {
                    WriteRaw(c switch { '\t' => "&#9;"u8, '\n' => "&#10;"u8, _ => "&#13;"u8 });
                }
                else
                {
                    WriteByte((byte)c);
                }

                return 1;
            case '_':
                // Literal text that already looks like an _xHHHH_ escape must have its underscore encoded.
                if (LooksLikeUnderscoreEscape(rest))
                {
                    WriteRaw("_x005F_"u8);
                    return 1;
                }

                WriteByte((byte)'_');
                return 1;
            default:
                WriteInvalidCharEscape(c);
                return 1;
        }
    }

    private static bool LooksLikeUnderscoreEscape(scoped ReadOnlySpan<char> text) =>
        text.Length >= 7 &&
        (text[1] is 'x' or 'X') &&
        char.IsAsciiHexDigit(text[2]) && char.IsAsciiHexDigit(text[3]) &&
        char.IsAsciiHexDigit(text[4]) && char.IsAsciiHexDigit(text[5]) &&
        text[6] == '_';

    /// <summary>Encodes an XML-invalid control character using the OOXML <c>_xHHHH_</c> convention.</summary>
    private void WriteInvalidCharEscape(char c)
    {
        EnsureCapacity(7);
        Span<byte> span = _buffer.AsSpan(_position);
        "_x"u8.CopyTo(span);
        ((int)c).TryFormat(span[2..], out _, "X4", CultureInfo.InvariantCulture);
        span[6] = (byte)'_';
        _position += 7;
    }

    private void TranscodeChunk(scoped ReadOnlySpan<char> text)
    {
        while (!text.IsEmpty)
        {
            EnsureCapacity(Math.Min(text.Length * 3, 1024));
            Utf8.FromUtf16(text, _buffer.AsSpan(_position), out int charsRead, out int bytesWritten,
                replaceInvalidSequences: true, isFinalBlock: true);
            _position += bytesWritten;
            text = text[charsRead..];
        }
    }

    private void WriteByte(byte value)
    {
        EnsureCapacity(1);
        _buffer[_position++] = value;
    }

    private void EnsureCapacity(int needed)
    {
        if (_buffer.Length - _position >= needed)
            return;

        byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_buffer.Length * 2, _position + needed));
        _buffer.AsSpan(0, _position).CopyTo(bigger);
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = bigger;
    }
}
