using System.Buffers.Binary;
using System.Text;

namespace Cellwright.Xls.Biff;

/// <summary>
/// Reading the string shapes of BIFF. Every one of them is the same idea — a length, a flags
/// byte, then characters that are either one byte each or two — wrapped differently depending
/// on where in the format it appears.
/// </summary>
/// <remarks>
/// The awkward part is the shared-string table. A string there may straddle a <c>Continue</c>
/// boundary, and the flags byte is repeated at the start of each piece, so the halves can be
/// encoded differently. <see cref="SstCursor"/> exists for that one case; everywhere else a
/// string is wholly inside its record.
/// </remarks>
internal static class BiffStrings
{
    /// <summary>
    /// Reads an <c>XLUnicodeString</c>: a two-byte character count, a flags byte, then the
    /// characters. Advances <paramref name="offset"/> past the whole thing.
    /// </summary>
    public static string ReadUnicodeString(
        ReadOnlySpan<byte> body,
        ref int offset,
        DocumentLoadBudgetState? loadBudget = null)
    {
        if (offset + 3 > body.Length)
        {
            offset = body.Length;
            return "";
        }

        int characters = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += 2;
        return ReadCharacters(body, ref offset, characters, loadBudget);
    }

    /// <summary>
    /// Reads a <c>ShortXLUnicodeString</c>, whose count is a single byte. Sheet names and font
    /// names use this form.
    /// </summary>
    public static string ReadShortUnicodeString(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 2 > body.Length)
        {
            offset = body.Length;
            return "";
        }

        int characters = body[offset];
        offset += 1;
        return ReadCharacters(body, ref offset, characters, loadBudget: null);
    }

    /// <summary>
    /// Reads the flags byte and the characters after a count has already been consumed. This is
    /// the body every unicode string in the format shares.
    /// </summary>
    private static string ReadCharacters(
        ReadOnlySpan<byte> body,
        ref int offset,
        int characters,
        DocumentLoadBudgetState? loadBudget)
    {
        if (offset >= body.Length)
            return "";

        byte flags = body[offset++];
        bool wide = (flags & 0x01) != 0;
        bool hasFarEast = (flags & 0x04) != 0;
        bool hasRuns = (flags & 0x08) != 0;

        int runCount = 0;
        int farEastSize = 0;
        if (hasRuns && offset + 2 <= body.Length)
        {
            runCount = BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
            offset += 2;
        }

        if (hasFarEast && offset + 4 <= body.Length)
        {
            farEastSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
            offset += 4;
        }

        int byteCount = wide ? characters * 2 : characters;
        byteCount = Math.Min(byteCount, Math.Max(0, body.Length - offset));
        loadBudget?.EnsureCellTextLength(wide ? byteCount / 2 : byteCount);
        string text = Decode(body.Slice(offset, byteCount), wide);
        offset += byteCount;

        // The formatting runs and the phonetic block trail the text. Neither has a place in the
        // model, so they are stepped over rather than parsed.
        offset += runCount * 4;
        offset += farEastSize;
        offset = Math.Min(offset, body.Length);
        return text;
    }

    /// <summary>
    /// Reads a byte string of BIFF2-5: a length and then characters in the workbook's code page.
    /// </summary>
    /// <param name="body">Record body.</param>
    /// <param name="offset">Where the string starts; advanced past it.</param>
    /// <param name="encoding">Code page the file declared.</param>
    /// <param name="lengthBytes">Whether the length is one byte or two.</param>
    public static string ReadByteString(ReadOnlySpan<byte> body, ref int offset, Encoding encoding, int lengthBytes)
    {
        if (offset + lengthBytes > body.Length)
        {
            offset = body.Length;
            return "";
        }

        int length = lengthBytes == 1
            ? body[offset]
            : BinaryPrimitives.ReadUInt16LittleEndian(body[offset..]);
        offset += lengthBytes;

        length = Math.Min(length, Math.Max(0, body.Length - offset));
        string text = encoding.GetString(body.Slice(offset, length));
        offset += length;
        return text;
    }

    /// <summary>
    /// Decodes characters that are either UTF-16 or the compressed form, where each byte is the
    /// low half of a UTF-16 unit. The compressed form is Latin-1, and not by coincidence: the
    /// two agree exactly over the whole byte range.
    /// </summary>
    private static string Decode(ReadOnlySpan<byte> bytes, bool wide) =>
        wide ? Encoding.Unicode.GetString(bytes) : Encoding.Latin1.GetString(bytes);

    /// <summary>
    /// Walks the strings of a shared-string table, which is the one place a string may be split
    /// across records. Each piece restates whether it is wide, so the two halves of one string
    /// can be encoded differently — hence a cursor rather than a plain offset.
    /// </summary>
    public ref struct SstCursor
    {
        private readonly ReadOnlySpan<byte> _body;
        private readonly IReadOnlyList<int> _pieces;
        private readonly DocumentLoadBudgetState? _loadBudget;
        private int _offset;
        private int _piece;
        private int _pieceEnd;

        /// <summary>Creates a cursor over an assembled SST record.</summary>
        /// <param name="body">The joined record body.</param>
        /// <param name="pieceSizes">Sizes of the pieces it arrived in, in order.</param>
        /// <param name="offset">Where to start reading, past the record's own header fields.</param>
        /// <param name="loadBudget">Shared counters used to reject oversized strings before decoding.</param>
        public SstCursor(
            ReadOnlySpan<byte> body,
            IReadOnlyList<int> pieceSizes,
            int offset,
            DocumentLoadBudgetState? loadBudget = null)
        {
            _body = body;
            _pieces = pieceSizes;
            _loadBudget = loadBudget;
            _offset = offset;
            _piece = 0;
            _pieceEnd = pieceSizes.Count > 0 ? pieceSizes[0] : body.Length;
            while (_pieceEnd <= offset && _piece + 1 < _pieces.Count)
                _pieceEnd += _pieces[++_piece];
        }

        /// <summary>Whether anything is left to read.</summary>
        public readonly bool HasMore => _offset < _body.Length;

        /// <summary>Reads the next string of the table.</summary>
        public string ReadNext() => ReadNext(out _);

        /// <summary>Reads the next string of the table and any character-formatting runs attached to it.</summary>
        public string ReadNext(out (int CharPos, int FontIndex)[]? runs)
        {
            runs = null;

            // A string may end exactly on a record boundary, leaving the cursor at the seam.
            // The next record then opens with a whole string, not with a repeated flags byte.
            while (_offset >= _pieceEnd && _piece + 1 < _pieces.Count)
                _pieceEnd += _pieces[++_piece];

            if (_offset + 3 > _body.Length)
            {
                _offset = _body.Length;
                return "";
            }

            int characters = BinaryPrimitives.ReadUInt16LittleEndian(_body[_offset..]);
            _offset += 2;
            byte flags = _body[_offset++];
            bool wide = (flags & 0x01) != 0;
            bool hasFarEast = (flags & 0x04) != 0;
            bool hasRuns = (flags & 0x08) != 0;

            int runCount = 0;
            int farEastSize = 0;
            if (hasRuns && _offset + 2 <= _body.Length)
            {
                runCount = BinaryPrimitives.ReadUInt16LittleEndian(_body[_offset..]);
                _offset += 2;
            }

            if (hasFarEast && _offset + 4 <= _body.Length)
            {
                farEastSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(_body[_offset..]);
                _offset += 4;
            }

            _loadBudget?.EnsureCellTextLength(characters);
            string text = ReadSplitCharacters(characters, wide);
            runs = ReadRuns(runCount);
            Skip(farEastSize);
            return text;
        }

        private (int CharPos, int FontIndex)[]? ReadRuns(int runCount)
        {
            if (runCount <= 0)
                return null;

            int availableRuns = Math.Min(runCount, (_body.Length - _offset) / 4);
            if (availableRuns <= 0)
            {
                Skip(runCount * 4);
                return null;
            }

            var runs = new (int CharPos, int FontIndex)[availableRuns];
            for (int i = 0; i < availableRuns; i++)
            {
                while (_offset >= _pieceEnd && _piece + 1 < _pieces.Count)
                    _pieceEnd += _pieces[++_piece];

                int charPos = BinaryPrimitives.ReadUInt16LittleEndian(_body[_offset..]);
                int fontIndex = BinaryPrimitives.ReadUInt16LittleEndian(_body[(_offset + 2)..]);
                _offset += 4;
                runs[i] = (charPos, fontIndex);
            }

            if (runCount > availableRuns)
                Skip((runCount - availableRuns) * 4);

            while (_offset >= _pieceEnd && _piece + 1 < _pieces.Count)
                _pieceEnd += _pieces[++_piece];

            return runs;
        }

        /// <summary>
        /// Reads the characters, crossing piece boundaries. At each boundary the format repeats
        /// the flags byte, and it may change: a string can start compressed and continue wide.
        /// </summary>
        private string ReadSplitCharacters(int characters, bool wide)
        {
            StringBuilder? builder = null;
            string? single = null;

            while (characters > 0 && _offset < _body.Length)
            {
                int available = Math.Max(0, CurrentPieceEnd() - _offset);
                if (available == 0)
                {
                    if (!AdvancePiece())
                        break;

                    // Every continuation restates the encoding of what follows.
                    if (_offset < _body.Length)
                        wide = (_body[_offset++] & 0x01) != 0;
                    continue;
                }

                int width = wide ? 2 : 1;
                int take = Math.Min(characters, available / width);
                if (take == 0)
                {
                    // An odd byte left before the boundary cannot hold a wide character.
                    if (!AdvancePiece())
                        break;
                    if (_offset < _body.Length)
                        wide = (_body[_offset++] & 0x01) != 0;
                    continue;
                }

                string chunk = Decode(_body.Slice(_offset, take * width), wide);
                _offset += take * width;
                characters -= take;

                if (single is null && builder is null)
                {
                    single = chunk;
                }
                else
                {
                    builder ??= new StringBuilder(single);
                    single = null;
                    builder.Append(chunk);
                }
            }

            return builder?.ToString() ?? single ?? "";
        }

        private readonly int CurrentPieceEnd() => _pieces.Count == 0 ? _body.Length : _pieceEnd;

        private bool AdvancePiece()
        {
            if (_piece + 1 >= _pieces.Count)
                return false;

            _offset = _pieceEnd;
            _pieceEnd += _pieces[++_piece];
            return true;
        }

        private void Skip(int bytes)
        {
            _offset = Math.Min(_offset + bytes, _body.Length);
            while (_offset >= _pieceEnd && _piece + 1 < _pieces.Count)
                _pieceEnd += _pieces[++_piece];
        }
    }
}
