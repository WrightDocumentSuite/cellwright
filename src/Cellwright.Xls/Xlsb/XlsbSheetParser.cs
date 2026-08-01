using System.Buffers.Binary;
using Cellwright.Model;
using Cellwright.Streaming;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads one worksheet part into the same <see cref="WorksheetContentSink"/> the xlsx and xls
/// readers fill.
/// </summary>
/// <remarks>
/// One thing here is unlike every other reader in this library: a cell record does not say
/// which row it is in. Every cell carries only its column, and the row is whichever
/// <c>BrtRowHdr</c> came before it (MS-XLSB 2.5.10). The parser therefore has a current row,
/// and a part that opens with a cell rather than a row header has nowhere to put it.
/// </remarks>
internal sealed class XlsbSheetParser(
    XlsGlobals globals,
    PtgDecoder decoder,
    int sheetIndex,
    Action<WorkbookWarning> warn)
{
    private int _row;
    private bool _hasRow;

    /// <summary>Shared formula masters by the range they cover, which members fall inside.</summary>
    private readonly List<(RangeRef Range, byte[] Rgce)> _sharedFormulas = [];

    /// <summary>Cells whose text came from a formula, waiting for the decode pass at the end.</summary>
    private readonly List<(CellRef Cell, byte[] Rgce)> _formulas = [];

    /// <summary>Cells that hold no formula of their own but fall inside a shared one.</summary>
    private readonly List<CellRef> _sharedMembers = [];

    /// <summary>Collects conditional formatting: a range header followed by its rules.</summary>
    private readonly XlsbConditionalFormats _conditionalFormats = new(decoder);

    public void Parse(byte[] part, WorksheetContentSink sink)
    {
        var reader = new Biff12Reader(part);
        while (reader.Read())
        {
            ReadOnlySpan<byte> body = reader.Record;
            switch (reader.RecordType)
            {
                case Biff12RecordType.RowHdr:
                    ReadRow(sink, body);
                    break;

                case Biff12RecordType.CellBlank:
                    SetCell(sink, body, CellValue.Blank);
                    break;
                case Biff12RecordType.CellRk:
                    ReadRk(sink, body);
                    break;
                case Biff12RecordType.CellError:
                    ReadError(sink, body);
                    break;
                case Biff12RecordType.CellBool:
                    ReadBool(sink, body);
                    break;
                case Biff12RecordType.CellReal:
                    ReadReal(sink, body);
                    break;
                case Biff12RecordType.CellSt:
                    ReadInlineString(sink, body);
                    break;
                case Biff12RecordType.CellIsst:
                    ReadSharedString(sink, body);
                    break;
                case Biff12RecordType.CellRString:
                    ReadRichString(sink, body);
                    break;

                case Biff12RecordType.FmlaNum or Biff12RecordType.FmlaString or
                     Biff12RecordType.FmlaBool or Biff12RecordType.FmlaError:
                    ReadFormulaCell(sink, body, reader.RecordType);
                    break;

                case Biff12RecordType.ShrFmla:
                    ReadSharedFormula(body, tokensAt: 16);
                    break;

                // An array formula covers a range the same way, and its members point at it
                // with the same token; only the byte between the range and the tokens differs.
                case Biff12RecordType.ArrFmla:
                    ReadSharedFormula(body, tokensAt: 17);
                    break;

                case Biff12RecordType.ColInfo:
                    ReadColumns(sink, body);
                    break;
                case Biff12RecordType.MergeCell:
                    ReadMerge(sink, body);
                    break;
                case Biff12RecordType.HLink:
                    ReadHyperlink(sink, body);
                    break;

                case Biff12RecordType.DVal:
                    XlsbValidationRecord.Read(body, sink, decoder, warn);
                    break;

                case Biff12RecordType.BeginConditionalFormatting:
                    _conditionalFormats.BeginGroup(body);
                    break;

                case Biff12RecordType.BeginCFRule:
                    _conditionalFormats.ReadRule(body);
                    break;

                case Biff12RecordType.EndCFRule:
                    _conditionalFormats.EndRule();
                    break;

                case Biff12RecordType.BeginColorScale:
                    _conditionalFormats.BeginColorScale();
                    break;

                case Biff12RecordType.BeginDatabar:
                    _conditionalFormats.BeginDataBar(body);
                    break;

                case Biff12RecordType.BeginIconSet:
                    _conditionalFormats.BeginIconSet(body);
                    break;

                case Biff12RecordType.CFVO:
                    _conditionalFormats.ReadValueObject(body);
                    break;

                case Biff12RecordType.Color:
                    _conditionalFormats.ReadColor(body);
                    break;

                default:
                    break;
            }
        }

        Finish(sink);
    }

    private void ReadRow(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 13)
            return;

        _row = (int)BinaryPrimitives.ReadUInt32LittleEndian(body) + 1;
        _hasRow = SheetLimits.IsValidRow(_row);
        if (!_hasRow)
            return;

        int styleIndex = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]);
        int twips = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);

        // Byte 10 carries only the two padding hints; everything that matters about the row is
        // in byte 11 (MS-XLSB 2.4.770).
        byte padding = body[10];
        byte flags = body[11];

        bool customHeight = (flags & 0x20) != 0;
        bool hasStyle = (flags & 0x40) != 0;
        var definition = new RowDefinition(
            Row: _row,
            Height: customHeight ? twips / 20.0 : null,
            Hidden: (flags & 0x10) != 0,
            OutlineLevel: flags & 0x07,
            Collapsed: (flags & 0x08) != 0,
            RawStyleIndex: 0,
            ThickTop: (padding & 0x01) != 0,
            ThickBottom: (padding & 0x02) != 0);

        if (definition.IsEmpty && !hasStyle)
            return;

        sink.OnRow(definition);
        if (hasStyle)
            sink.Sheet.RowTable[_row] = sink.Sheet.RowTable[_row] with { Style = globals.StyleOf(styleIndex) };
    }

    private void ReadRk(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 12)
            return;

        double number = BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(body[8..]));
        SetCell(sink, body, number);
    }

    private void ReadReal(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 16)
            return;

        SetCell(sink, body, BinaryPrimitives.ReadDoubleLittleEndian(body[8..]));
    }

    private void ReadBool(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 9)
            return;

        SetCell(sink, body, CellValue.Boolean(body[8] != 0));
    }

    private void ReadError(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 9)
            return;

        SetCell(sink, body, CellValue.Error(Biff12Values.ReadError(body[8])));
    }

    private void ReadInlineString(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        int offset = 8;
        SetCell(sink, body, Text(Biff12Values.ReadString(body, ref offset)));
    }

    /// <summary>Rich text written straight into the cell; the runs after it are not modelled.</summary>
    private void ReadRichString(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 9)
            return;

        int offset = 9;
        SetCell(sink, body, Text(Biff12Values.ReadString(body, ref offset)));
    }

    private void ReadSharedString(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 12)
            return;

        int index = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[8..]);
        SetCell(sink, body, Text(globals.StringAt(index)));
    }

    /// <summary>
    /// A formula cell: its cached result, then the tokens. A cell inside a shared or array
    /// formula does not repeat the expression; it is matched to its master once the sheet has
    /// been read.
    /// </summary>
    private void ReadFormulaCell(WorksheetContentSink sink, ReadOnlySpan<byte> body, int type)
    {
        if (!TryCell(body, out CellRef cell, out int styleIndex))
            return;

        int offset = 8;
        if (type == Biff12RecordType.FmlaNum && offset + 8 <= body.Length)
        {
            SetCell(sink, body, BinaryPrimitives.ReadDoubleLittleEndian(body[offset..]));
            offset += 8;
        }
        else
        {
            sink.OnPooledCell(cell, ReadCachedValue(body, ref offset, type), globals.StyleOf(styleIndex), formula: null);
        }

        if (offset + 6 > body.Length)
            return;

        offset += 2;                                        // The flags of the formula itself.
        int tokens = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;

        if (tokens <= 0 || offset + tokens > body.Length)
        {
            _sharedMembers.Add(cell);
            return;
        }

        // A member of a shared or array formula does not repeat the expression: it holds a
        // single PtgExp naming the row the master sits in (MS-XLSB 2.5.98.40). Decoding that
        // token is impossible on its own, so the cell waits for the master instead.
        ReadOnlySpan<byte> rgce = body.Slice(offset, tokens);
        if (rgce[0] == PtgExp)
            _sharedMembers.Add(cell);
        else
            _formulas.Add((cell, rgce.ToArray()));
    }

    /// <summary>The token that stands in for a shared or array formula stated elsewhere.</summary>
    private const byte PtgExp = 0x01;

    private CellValue ReadCachedValue(ReadOnlySpan<byte> body, ref int offset, int type) => type switch
    {
        Biff12RecordType.FmlaBool when offset < body.Length => CellValue.Boolean(body[offset++] != 0),
        Biff12RecordType.FmlaError when offset < body.Length => CellValue.Error(Biff12Values.ReadError(body[offset++])),
        Biff12RecordType.FmlaString => Text(Biff12Values.ReadString(body, ref offset)),
        _ => CellValue.Blank,
    };

    /// <summary>
    /// A shared or array formula: the range it covers and the expression written from its
    /// top-left cell.
    /// </summary>
    /// <param name="body">The record body.</param>
    /// <param name="tokensAt">Where the token count sits, which is the only difference between the two.</param>
    private void ReadSharedFormula(ReadOnlySpan<byte> body, int tokensAt)
    {
        if (body.Length < tokensAt + 4)
            return;

        RangeRef range = ReadRange(body);
        int tokens = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[tokensAt..]);
        int start = tokensAt + 4;
        if (tokens <= 0 || start + tokens > body.Length)
            return;

        _sharedFormulas.Add((range, body.Slice(start, tokens).ToArray()));
    }

    private void ReadColumns(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 18)
            return;

        int first = (int)BinaryPrimitives.ReadUInt32LittleEndian(body) + 1;
        int last = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]) + 1;
        int width = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[8..]);
        int styleIndex = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[12..]);
        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body[16..]);

        last = Math.Min(last, SheetLimits.MaxColumn);
        for (int column = Math.Max(first, 1); column <= last; column++)
        {
            sink.OnColumn(new ColumnDefinition(
                Column: column,
                Width: width / 256.0,
                Hidden: (flags & 0x01) != 0,
                OutlineLevel: (flags >> 8) & 0x07,
                Collapsed: (flags & 0x1000) != 0,
                RawStyleIndex: 0,
                BestFit: (flags & 0x04) != 0));

            if (styleIndex != 0)
                sink.Sheet.ColumnTable[column] = sink.Sheet.ColumnTable[column] with { Style = globals.StyleOf(styleIndex) };
        }
    }

    private static void ReadMerge(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 16)
            return;

        sink.OnMerge(ReadRange(body));
    }

    private static void ReadHyperlink(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 16)
            return;

        RangeRef range = ReadRange(body);
        int offset = 16;
        string? relationship = Biff12Values.ReadNullableString(body, ref offset);
        string location = Biff12Values.ReadString(body, ref offset);
        string tooltip = Biff12Values.ReadString(body, ref offset);
        string display = Biff12Values.ReadString(body, ref offset);

        sink.OnHyperlink(
            range,
            string.IsNullOrEmpty(relationship) ? null : relationship,
            location.Length == 0 ? null : location,
            tooltip.Length == 0 ? null : tooltip,
            display.Length == 0 ? null : display);
    }

    /// <summary>
    /// Decodes the formulas once the whole sheet is known. A cell inside a shared group has no
    /// tokens of its own, so it borrows the master's and decodes them from where it sits.
    /// </summary>
    private void Finish(WorksheetContentSink sink)
    {
        foreach ((CellRef cell, byte[] rgce) in _formulas)
            Render(sink, cell, rgce);

        foreach (CellRef cell in _sharedMembers)
        {
            foreach ((RangeRef range, byte[] rgce) in _sharedFormulas)
            {
                if (range.Contains(cell))
                {
                    Render(sink, cell, rgce);
                    break;
                }
            }
        }

        _conditionalFormats.Finish(sink.Sheet, warn);

        _formulas.Clear();
        _sharedMembers.Clear();
        _sharedFormulas.Clear();
    }

    private void Render(WorksheetContentSink sink, CellRef cell, byte[] rgce)
    {
        string? text = decoder.Decode(rgce, cell, sheetIndex);
        if (text is null)
        {
            warn(new WorkbookWarning(
                "UndecodedFormula",
                "The formula holds something this reader cannot render; its last value was kept.",
                cell));
            return;
        }

        sink.Sheet.Formulas.Set(cell.Row, cell.Column, text);
    }

    private static RangeRef ReadRange(ReadOnlySpan<byte> body) => new(
        new CellRef((int)BinaryPrimitives.ReadUInt32LittleEndian(body) + 1,
            (int)BinaryPrimitives.ReadUInt32LittleEndian(body[8..]) + 1),
        new CellRef((int)BinaryPrimitives.ReadUInt32LittleEndian(body[4..]) + 1,
            (int)BinaryPrimitives.ReadUInt32LittleEndian(body[12..]) + 1));

    /// <summary>
    /// Resolves which cell a record is about. The column and the style come from the record's
    /// own <c>Cell</c> structure; the row is whichever header the parser last passed.
    /// </summary>
    private bool TryCell(ReadOnlySpan<byte> body, out CellRef cell, out int styleIndex)
    {
        cell = default;
        if (!_hasRow || !Biff12Values.TryReadCell(body, out int column, out styleIndex))
        {
            styleIndex = 0;
            return false;
        }

        column++;
        if (!SheetLimits.IsValidColumn(column))
            return false;

        cell = new CellRef(_row, column);
        return true;
    }

    /// <summary>
    /// Writes a number as the value its format says it is: a date in a spreadsheet is a number
    /// with a date format on it and nothing else, so the format table is the only clue.
    /// </summary>
    private void SetCell(WorksheetContentSink sink, ReadOnlySpan<byte> body, double number)
    {
        if (!TryCell(body, out CellRef cell, out int styleIndex))
            return;

        CellValue value = globals.TemporalOf(styleIndex) switch
        {
            Styles.TemporalKind.Date => TemporalValue.AsDate(number, globals.DateSystem),
            Styles.TemporalKind.Duration => TemporalValue.AsDuration(number),
            _ => CellValue.Number(number),
        };

        sink.OnPooledCell(cell, value, globals.StyleOf(styleIndex), formula: null);
    }

    private void SetCell(WorksheetContentSink sink, ReadOnlySpan<byte> body, in CellValue value)
    {
        if (TryCell(body, out CellRef cell, out int styleIndex))
            sink.OnPooledCell(cell, value, globals.StyleOf(styleIndex), formula: null);
    }

    /// <summary>Cuts text to what a cell can hold, saying so rather than dropping the cell.</summary>
    private CellValue Text(string text)
    {
        if (text.Length <= SheetLimits.MaxCellTextLength)
            return CellValue.Text(text);

        warn(new WorkbookWarning(
            "TextTooLong",
            $"A cell holds {text.Length} characters, more than the {SheetLimits.MaxCellTextLength} a cell can carry; it was cut."));
        return CellValue.Text(text[..SheetLimits.MaxCellTextLength]);
    }
}
