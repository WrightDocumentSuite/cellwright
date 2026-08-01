using System.Buffers.Binary;
using System.Text;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Gathers the three records that together hold one cell note. The anchor and the author live
/// in a <c>Note</c> record; the text lives in a <c>TxO</c> record of a drawing object; and the
/// <c>Obj</c> record between them carries the object id that connects the two. Notes usually
/// trail the drawing records by half a substream, so the pieces are collected while the sheet
/// is read and joined at its end.
/// </summary>
internal sealed class CommentCollector
{
    /// <summary>Object type of a comment in <c>ftCmo.ot</c> (MS-XLS FtCmo).</summary>
    private const ushort NoteObjectType = 0x0019;

    /// <summary>Comment text by drawing object id, captured from the TxO after each note's Obj.</summary>
    private readonly Dictionary<int, string> _texts = [];

    /// <summary>Notes waiting for their text: where they sit, who wrote them, what they point at.</summary>
    private readonly List<(CellRef Cell, string Author, int ObjectId)> _notes = [];

    /// <summary>Id of the last note object seen; the next TxO record carries its text.</summary>
    private int _pendingObjectId = -1;

    /// <summary>
    /// Reads the <c>ftCmo</c> that opens every Obj record. Only a note's object id is kept, and
    /// any other object type clears the pending id: text boxes and controls have TxO records of
    /// their own, and without the reset their text would land on the previous note.
    /// </summary>
    public void OnObj(ReadOnlySpan<byte> body)
    {
        if (body.Length < 8 || BinaryPrimitives.ReadUInt16LittleEndian(body) != 0x0015)
        {
            _pendingObjectId = -1;
            return;
        }

        ushort objectType = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        ushort id = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        _pendingObjectId = objectType == NoteObjectType ? id : -1;
    }

    /// <summary>
    /// Captures the text of the TxO record that belongs to the pending note object. The record
    /// itself only says how long the text is; the characters arrive in the Continue records
    /// after it, each fragment restating its own encoding.
    /// </summary>
    /// <param name="body">The joined record body.</param>
    /// <param name="pieces">Sizes of the pieces it arrived in, or <see langword="null"/> when whole.</param>
    public void OnTxO(ReadOnlySpan<byte> body, IReadOnlyList<int>? pieces)
    {
        if (_pendingObjectId < 0)
            return;

        int id = _pendingObjectId;
        _pendingObjectId = -1;
        if (body.Length < 12)
            return;

        int characters = BinaryPrimitives.ReadUInt16LittleEndian(body[10..]);
        if (characters == 0)
        {
            _texts[id] = "";
            return;
        }

        // Most producers put the text in continuation records. A few write it straight after
        // the record's own fields, and then there is a single piece to read from.
        _texts[id] = pieces is { Count: >= 2 }
            ? ReadFragmentedText(body, pieces, characters)
            : ReadInlineText(body, characters);
    }

    /// <summary>Reads text that follows the record's fixed fields instead of arriving separately.</summary>
    private static string ReadInlineText(ReadOnlySpan<byte> body, int characters)
    {
        // The fields end at 18 bytes for a note, whose object type has no control block.
        const int TextStart = 18;
        if (TextStart >= body.Length)
            return "";

        int offset = TextStart;
        bool wide = (body[offset++] & 0x01) != 0;
        int take = Math.Min(characters, (body.Length - offset) / (wide ? 2 : 1));
        if (take <= 0)
            return "";

        ReadOnlySpan<byte> bytes = body.Slice(offset, take * (wide ? 2 : 1));
        return wide ? Encoding.Unicode.GetString(bytes) : Encoding.Latin1.GetString(bytes);
    }

    /// <summary>
    /// Reads the text fragments that follow the TxO record proper. Each fragment opens with its
    /// own flags byte, so a string can start compressed and continue wide; reading stops once
    /// the promised characters are in hand, which leaves the formatting runs behind unread.
    /// </summary>
    private static string ReadFragmentedText(ReadOnlySpan<byte> body, IReadOnlyList<int> pieces, int characters)
    {
        var text = new StringBuilder(Math.Min(characters, 1024));
        int offset = pieces[0];
        for (int piece = 1; piece < pieces.Count && characters > 0; piece++)
        {
            int end = Math.Min(offset + pieces[piece], body.Length);
            if (offset >= end)
                break;

            bool wide = (body[offset++] & 0x01) != 0;
            int width = wide ? 2 : 1;
            int take = Math.Min(characters, (end - offset) / width);
            if (take > 0)
            {
                ReadOnlySpan<byte> bytes = body.Slice(offset, take * width);
                text.Append(wide ? Encoding.Unicode.GetString(bytes) : Encoding.Latin1.GetString(bytes));
                characters -= take;
            }

            offset = end;
        }

        return text.ToString();
    }

    /// <summary>Reads a <c>Note</c> record — a NoteSh: anchor cell, flags, object id, author.</summary>
    public void OnNote(ReadOnlySpan<byte> body)
    {
        if (body.Length < 10)
            return;

        int row = BinaryPrimitives.ReadUInt16LittleEndian(body) + 1;
        int column = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        if (!SheetLimits.IsValidRow(row) || !SheetLimits.IsValidColumn(column))
            return;

        int objectId = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        int cursor = 8;
        string author = BiffStrings.ReadUnicodeString(body, ref cursor);
        _notes.Add((new CellRef(row, column), author, objectId));
    }

    /// <summary>
    /// Attaches the collected notes to the sheet. The public <c>Add</c> is deliberate: a BIFF
    /// file has no comment part to replay on save, so the model must own these notes for the
    /// xlsx writer to regenerate them.
    /// </summary>
    public void Finish(Worksheet sheet, Action<WorkbookWarning> warn)
    {
        foreach ((CellRef cell, string author, int objectId) in _notes)
        {
            if (!_texts.TryGetValue(objectId, out string? text))
            {
                warn(new WorkbookWarning(
                    "CommentTextMissing",
                    $"A note in '{sheet.Name}' points at a drawing object that has no text; the note was kept empty.",
                    cell));
                text = "";
            }

            sheet.Comments.Add(cell, text, author.Length > 0 ? author : null);
        }

        _notes.Clear();
        _texts.Clear();
        _pendingObjectId = -1;
    }
}
