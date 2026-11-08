using System.Buffers.Binary;
using Cellwright.Model;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the comments part of a binary workbook: the authors, then the notes that point at
/// them by position in that list.
/// </summary>
/// <remarks>
/// The binary format keeps notes where the modern one does — in a part of their own beside the
/// sheet — rather than in the drawing layer, which is where the older binary format hides them.
/// That makes this the simplest of the three readers to write and the only one whose shape the
/// xlsx side would recognise.
/// </remarks>
internal static class XlsbCommentsPart
{
    public static void Read(byte[] part, Worksheet sheet, Action<WorkbookWarning> warn)
    {
        var authors = new List<string>();
        var reader = new Biff12Reader(part);

        CellRef cell = default;
        int authorIndex = -1;
        bool inComment = false;

        while (reader.Read())
        {
            ReadOnlySpan<byte> body = reader.Record;
            switch (reader.RecordType)
            {
                case Biff12RecordType.CommentAuthor:
                    {
                        int offset = 0;
                        authors.Add(Biff12Values.ReadString(body, ref offset));
                        break;
                    }

                // The note states which cell it belongs to and who wrote it; its text arrives
                // in the record that follows.
                case Biff12RecordType.BeginComment when body.Length >= 20:
                    authorIndex = BinaryPrimitives.ReadInt32LittleEndian(body);
                    cell = new CellRef(
                        (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]) + 1,
                        (int)BinaryPrimitives.ReadUInt32LittleEndian(body[12..]) + 1);
                    inComment = SheetLimits.IsValidRow(cell.Row) && SheetLimits.IsValidColumn(cell.Column);
                    break;

                case Biff12RecordType.CommentText when inComment:
                    {
                        string text = XlsbStringTable.ReadRichString(body);
                        string? author = (uint)authorIndex < (uint)authors.Count ? authors[authorIndex] : null;
                        sheet.Comments.Add(cell, text, string.IsNullOrEmpty(author) ? null : author);
                        inComment = false;
                        break;
                    }

                case Biff12RecordType.EndComment when inComment:
                    warn(new WorkbookWarning(
                        "CommentTextMissing",
                        $"A note in '{sheet.Name}' carries no text record; it was kept empty.",
                        cell));
                    sheet.Comments.Add(cell, "", null);
                    inComment = false;
                    break;

                default:
                    break;
            }
        }
    }
}
