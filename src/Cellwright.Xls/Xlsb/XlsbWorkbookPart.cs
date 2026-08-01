using System.Buffers.Binary;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the workbook part: which sheets there are, the names defined over them, the date
/// system, and the table of supporting links that 3-D references are written against.
/// </summary>
/// <remarks>
/// What it fills is <see cref="XlsGlobals"/> — the same container the older reader fills — so
/// that the formula decoder serves both. The two formats state the same things about a link
/// table in different records, and this is where that difference stops.
/// </remarks>
internal static class XlsbWorkbookPart
{
    /// <summary>Relationship ids of the sheets, in tab order, parallel to <c>globals.Sheets</c>.</summary>
    public static List<string?> Read(byte[] part, XlsGlobals globals)
    {
        var relationships = new List<string?>();
        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            switch (reader.RecordType)
            {
                case Biff12RecordType.WbProp:
                    ReadProperties(reader.Record, globals);
                    break;

                case Biff12RecordType.BundleSh:
                    ReadSheet(reader.Record, globals, relationships);
                    break;

                case Biff12RecordType.Name:
                    ReadName(reader.Record, globals);
                    break;

                // Every supporting link is one record, and the Xti entries index into them in
                // the order they appear, so the order has to be kept even for the kinds this
                // reader cannot follow.
                case Biff12RecordType.SupSelf or Biff12RecordType.SupSame:
                    globals.SupBooks.Add(new XlsSupBook());
                    break;

                case Biff12RecordType.SupAddin:
                    globals.SupBooks.Add(new XlsSupBook { IsAddin = true });
                    break;

                case Biff12RecordType.SupBookSrc:
                    globals.SupBooks.Add(ReadExternalBook(reader.Record));
                    break;

                case Biff12RecordType.ExternSheet:
                    ReadExternSheets(reader.Record, globals);
                    break;

                default:
                    break;
            }
        }

        return relationships;
    }

    private static void ReadProperties(ReadOnlySpan<byte> body, XlsGlobals globals)
    {
        if (body.Length < 4)
            return;

        // f1904 is the lowest bit of the first flags word.
        bool is1904 = (BinaryPrimitives.ReadUInt32LittleEndian(body) & 0x1) != 0;
        globals.DateSystem = is1904 ? ExcelDateSystem.Date1904 : ExcelDateSystem.Date1900;
    }

    /// <summary>
    /// Reads one sheet of the workbook's list. The record is a visibility state, a tab
    /// identifier, the relationship naming the part that holds the sheet, and the tab name.
    /// </summary>
    /// <remarks>
    /// Files exist with an extra four bytes ahead of the tab identifier, and reading them at
    /// the documented offset yields a relationship id of one null character, which matches no
    /// relationship and silently loses every sheet in the workbook. The two shapes are told
    /// apart by arithmetic rather than by guessing: only the right header makes the two strings
    /// end exactly where the record does.
    /// </remarks>
    private static void ReadSheet(ReadOnlySpan<byte> body, XlsGlobals globals, List<string?> relationships)
    {
        if (!TryReadSheet(body, header: 8, out string? relationship, out string name) &&
            !TryReadSheet(body, header: 12, out relationship, out name))
        {
            return;
        }

        // A sheet with no part of its own is a module sheet, which holds macros rather than a
        // grid; it keeps its place in the tab order so that sheet indexes still line up.
        relationships.Add(relationship);
        globals.Sheets.Add(new XlsSheetEntry(
            name,
            Offset: 0,
            BinaryPrimitives.ReadUInt32LittleEndian(body) switch
            {
                1 => SheetVisibility.Hidden,
                2 => SheetVisibility.VeryHidden,
                _ => SheetVisibility.Visible,
            },
            relationship is null ? XlsSheetKind.Other : XlsSheetKind.Worksheet));
    }

    private static bool TryReadSheet(
        ReadOnlySpan<byte> body, int header, out string? relationship, out string name)
    {
        relationship = null;
        name = "";
        if (body.Length < header + 8)
            return false;

        int offset = header;
        if (!Biff12Values.TryReadString(body, ref offset, out relationship) ||
            !Biff12Values.TryReadString(body, ref offset, out string? tabName))
        {
            return false;
        }

        name = tabName ?? "";

        // The two strings run to the end of the record and no further; anything else means the
        // header was the wrong length and these are not the strings at all.
        return offset == body.Length && name.Length > 0;
    }

    /// <summary>
    /// Reads a defined name. Only the name, its scope and its expression are kept; the strings
    /// that follow describe a macro's menu entry and help topic, which the model cannot hold.
    /// </summary>
    private static void ReadName(ReadOnlySpan<byte> body, XlsGlobals globals)
    {
        if (body.Length < 9)
            return;

        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body);
        uint scope = BinaryPrimitives.ReadUInt32LittleEndian(body[5..]);   // Past the shortcut key.

        int offset = 9;
        string name = Biff12Values.ReadString(body, ref offset);
        if (name.Length == 0 || offset + 4 > body.Length)
            return;

        int tokens = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (tokens < 0 || offset + tokens > body.Length)
            return;

        globals.Names.Add(new XlsDefinedName(
            name,
            body.Slice(offset, tokens).ToArray(),
            scope == 0xFFFFFFFF ? 0 : (int)scope + 1,
            (flags & 0x1) != 0,
            IsBuiltIn: (flags & 0x20) != 0));
    }

    private static XlsSupBook ReadExternalBook(ReadOnlySpan<byte> body)
    {
        int offset = 0;
        string? path = Biff12Values.ReadNullableString(body, ref offset);
        return new XlsSupBook { Path = path ?? "" };
    }

    private static void ReadExternSheets(ReadOnlySpan<byte> body, XlsGlobals globals)
    {
        if (body.Length < 4)
            return;

        int count = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        for (int i = 0; i < count && 4 + ((i + 1) * 12) <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(4 + (i * 12), 12);
            globals.ExternSheets.Add((
                (int)BinaryPrimitives.ReadUInt32LittleEndian(entry),
                BinaryPrimitives.ReadInt32LittleEndian(entry[4..]),
                BinaryPrimitives.ReadInt32LittleEndian(entry[8..])));
        }
    }
}
