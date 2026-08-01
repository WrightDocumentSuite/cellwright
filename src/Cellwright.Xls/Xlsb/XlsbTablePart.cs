using System.Buffers.Binary;
using Cellwright.Model;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads a table part: the range a table covers, the name formulas refer to it by, and the
/// names of its columns (MS-XLSB 2.4.100 and 2.4.101).
/// </summary>
/// <remarks>
/// Tables are read for two reasons. One is the model, which has held them since the xlsx side
/// was written. The other is formulas: a structured reference is a single token in this format,
/// and rendering it as <c>Table[Column]</c> needs the names, which live only here.
/// </remarks>
internal static class XlsbTablePart
{
    /// <summary>Fixed part of <c>BrtBeginList</c> before its strings begin.</summary>
    private const int ListHeader = 64;

    /// <summary>Fixed part of <c>BrtBeginListCol</c> before its strings begin.</summary>
    private const int ColumnHeader = 24;

    /// <summary>Reads a table part into the sheet, and into the map formulas resolve against.</summary>
    /// <param name="part">Bytes of the table part.</param>
    /// <param name="sheet">Sheet the table sits on.</param>
    /// <param name="byId">
    /// Tables by the numeric identifier a <c>PtgList</c> token names them by. It is kept beside
    /// the model rather than in it: the number means nothing outside this one file.
    /// </param>
    public static void Read(byte[] part, Worksheet sheet, IDictionary<uint, ExcelTable> byId)
    {
        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            if (reader.RecordType != Biff12RecordType.BeginList)
                continue;

            if (ReadList(reader, out uint id) is { } table)
            {
                sheet.Tables.AddFromFile(table);
                byId[id] = table;
            }
        }
    }

    private static ExcelTable? ReadList(Biff12Reader reader, out uint id)
    {
        id = 0;
        ReadOnlySpan<byte> body = reader.Record;
        if (body.Length < ListHeader)
            return null;

        var range = new RangeRef(
            new CellRef(
                (int)BinaryPrimitives.ReadUInt32LittleEndian(body) + 1,
                (int)BinaryPrimitives.ReadUInt32LittleEndian(body[8..]) + 1),
            new CellRef(
                (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]) + 1,
                (int)BinaryPrimitives.ReadUInt32LittleEndian(body[12..]) + 1));

        id = BinaryPrimitives.ReadUInt32LittleEndian(body[20..]);
        bool hasHeader = BinaryPrimitives.ReadUInt32LittleEndian(body[24..]) != 0;
        bool hasTotals = BinaryPrimitives.ReadUInt32LittleEndian(body[28..]) != 0;

        int offset = ListHeader;
        string? name = Biff12Values.ReadNullableString(body, ref offset);
        string? displayName = Biff12Values.ReadNullableString(body, ref offset);

        // A formula names a table by its display name; the other one is for code that goes
        // looking for it. Either may be absent, and only the pair together is always present.
        string chosen = Pick(displayName, name);
        if (chosen.Length == 0 ||
            !SheetLimits.IsValidRow(range.End.Row) || !SheetLimits.IsValidColumn(range.End.Column))
        {
            return null;
        }

        List<string> columns = ReadColumns(reader, range.ColumnCount);
        if (columns.Count != range.ColumnCount)
            return null;

        return new ExcelTable
        {
            Name = chosen,
            Range = range,
            Columns = columns,
            HasHeaderRow = hasHeader,
            HasTotalsRow = hasTotals,
        };
    }

    /// <summary>
    /// Reads the column names that follow, stopping at the end of the table. A column whose
    /// name the file leaves out is given the same generated one Excel shows.
    /// </summary>
    private static List<string> ReadColumns(Biff12Reader reader, int expected)
    {
        var columns = new List<string>(expected);
        while (reader.Read())
        {
            if (reader.RecordType is Biff12RecordType.EndList or Biff12RecordType.BeginList)
                break;
            if (reader.RecordType != Biff12RecordType.BeginListCol)
                continue;

            ReadOnlySpan<byte> body = reader.Record;
            if (body.Length < ColumnHeader)
                continue;

            int offset = ColumnHeader;
            string? name = Biff12Values.ReadNullableString(body, ref offset);
            string? caption = Biff12Values.ReadNullableString(body, ref offset);

            // The caption is what the header cell says, and so what a structured reference
            // spells; the name is the programmatic identifier beside it.
            string chosen = Pick(caption, name);
            columns.Add(chosen.Length > 0 ? chosen : $"Column{columns.Count + 1}");
        }

        return columns;
    }

    private static string Pick(string? preferred, string? fallback) =>
        !string.IsNullOrEmpty(preferred) ? preferred : fallback ?? "";
}
