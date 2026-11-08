using System.Buffers.Binary;
using Cellwright.Model;
using Cellwright.Streaming;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads one sheet's substream — the cell table and everything around it — into a
/// <see cref="WorksheetContentSink"/>, the same receiver the xlsx reader fills.
/// </summary>
/// <param name="reader">Reader positioned just past the sheet's opening BOF.</param>
/// <param name="globals">What the workbook's global substream said.</param>
/// <param name="decoder">Renders the sheet's formulas.</param>
/// <param name="sheetIndex">Position of the sheet among those the file declares.</param>
/// <param name="substreamEnd">
/// Where the next substream begins, which is as far as this one can possibly run. A sheet may
/// hold substreams of its own — an embedded chart is one — and telling those apart from the
/// next sheet is only possible by knowing where the next sheet starts.
/// </param>
/// <param name="warn">Reports what could be recovered but not understood.</param>
/// <param name="loadBudget">Shared resource counters for this import.</param>
/// <param name="cancellationToken">Cancels record parsing.</param>
internal sealed class SheetParser(
    BiffReader reader,
    XlsGlobals globals,
    PtgDecoder decoder,
    int sheetIndex,
    int substreamEnd,
    Action<WorkbookWarning> warn,
    DocumentLoadBudgetState? loadBudget = null,
    CancellationToken cancellationToken = default)
{
    /// <summary>The pending formula, waiting for the String record that may follow it.</summary>
    private CellRef _pendingStringCell;
    private bool _hasPendingString;

    /// <summary>Shared formula masters by their top-left cell, which PtgExp points at.</summary>
    private readonly Dictionary<CellRef, byte[]> _sharedFormulas = [];

    /// <summary>Cells that hold a shared formula, resolved once the master turns up.</summary>
    private readonly List<(CellRef Cell, CellRef Master)> _sharedMembers = [];

    private readonly List<(CellRef Cell, byte[] Rgce)> _deferredFormulas = [];

    /// <summary>
    /// Collects the notes of the sheet, or <see langword="null"/> before BIFF8: the older Note
    /// record carries its text itself in a shape this reader does not take apart.
    /// </summary>
    private readonly CommentCollector? _comments =
        globals.Version >= BiffVersion.Biff8 ? new CommentCollector() : null;

    /// <summary>Collects the print settings, which arrive scattered over the whole substream.</summary>
    private readonly PageSetupCollector _pageSetup = new();

    /// <summary>The sheet's drawing layer, in the pieces it arrives in. See <see cref="EscherPictures.Join"/>.</summary>
    private readonly List<byte[]> _drawing = [];

    /// <summary>Collects conditional formatting: a range header followed by its rules.</summary>
    private readonly ConditionalFormatCollector _conditionalFormats = Attach(globals);

    private static ConditionalFormatCollector Attach(XlsGlobals globals)
    {
        var collector = new ConditionalFormatCollector();
        collector.Attach(globals);
        return collector;
    }

    /// <summary>Reads to the sheet's EOF, filling the sink.</summary>
    public void Parse(WorksheetContentSink sink)
    {
        Worksheet sheet = sink.Sheet;
        int frozenRows = 0;
        int frozenColumns = 0;
        bool hasFrozenPane = false;

        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            ReadOnlySpan<byte> body = reader.Record.Span;
            switch (reader.RecordType)
            {
                case BiffRecordType.Eof:
                    Finish(sink);
                    if (hasFrozenPane)
                        sink.OnFrozenPane(frozenRows, frozenColumns);
                    return;

                case BiffRecordType.Bof:
                    // A sheet can hold substreams of its own: an embedded chart is one, and
                    // everything after it — the window, the merges, the notes — still belongs
                    // to this sheet. Only a BOF where the next sheet starts ends this one.
                    if (reader.RecordOffset < substreamEnd)
                    {
                        SkipNestedSubstream();
                        break;
                    }

                    // A substream should close with EOF, and files exist where one does not.
                    warn(new WorkbookWarning(
                        "MissingEof",
                        $"The substream of '{sheet.Name}' has no end marker; it was taken to end where the next one begins."));
                    Finish(sink);
                    if (hasFrozenPane)
                        sink.OnFrozenPane(frozenRows, frozenColumns);
                    return;

                case BiffRecordType.Row:
                    ReadRow(sink, body);
                    break;

                case BiffRecordType.Blank:
                    ReadBlank(sink, body);
                    break;

                case BiffRecordType.MulBlank:
                    ReadMulBlank(sink, body);
                    break;

                case BiffRecordType.Rk:
                    ReadRk(sink, body);
                    break;

                case BiffRecordType.MulRk:
                    ReadMulRk(sink, body);
                    break;

                case BiffRecordType.Number:
                    ReadNumber(sink, body);
                    break;

                case BiffRecordType.BoolErr:
                    ReadBoolErr(sink, body);
                    break;

                case BiffRecordType.LabelSst:
                    ReadLabelSst(sink, body);
                    break;

                case BiffRecordType.Label or BiffRecordType.RString:
                    ReadLabel(sink, body);
                    break;

                case BiffRecordType.Formula:
                    ReadFormula(sink, body);
                    break;

                case BiffRecordType.String:
                    ReadPendingString(sink, body);
                    break;

                case BiffRecordType.ShrFmla:
                    ReadSharedFormula(body);
                    break;

                case BiffRecordType.Array:
                    ReadArrayFormula(sink, body);
                    break;

                case BiffRecordType.ColInfo:
                    ReadColumnInfo(sink, body);
                    break;

                // DefColWidth and DefaultRowHeight state the sheet-wide defaults, which the
                // model does not carry: a row or column keeps a size only when it has its own.

                case BiffRecordType.MergeCells:
                    ReadMergeCells(sink, body);
                    break;

                case BiffRecordType.Window2:
                    ReadWindow2(sink, body);
                    break;

                case BiffRecordType.Scl:
                    ReadZoom(sink, body);
                    break;

                case BiffRecordType.Pane:
                    ReadPane(body, ref frozenRows, ref frozenColumns, ref hasFrozenPane);
                    break;

                case BiffRecordType.HLink:
                    HyperlinkRecord.Read(body, sink, warn);
                    break;

                // The tooltip is a record of its own that follows the link it belongs to.
                case BiffRecordType.HLinkTooltip:
                    HyperlinkTooltipRecord.Read(body, sink.Sheet);
                    break;

                case BiffRecordType.Dv:
                    DataValidationRecord.Read(body, sink, decoder, warn);
                    break;

                case BiffRecordType.SheetExt:
                    ReadSheetExt(sink, body);
                    break;

                case BiffRecordType.Protect:
                    if (body.Length >= 2 && BinaryPrimitives.ReadUInt16LittleEndian(body) != 0)
                        sheet.Protection.IsEnabled = true;
                    break;

                case BiffRecordType.MsoDrawing:
                    _drawing.Add(body.ToArray());
                    break;

                case BiffRecordType.Obj:
                    _comments?.OnObj(body);
                    break;

                case BiffRecordType.TxO:
                    _comments?.OnTxO(body, reader.WasContinued ? reader.ContinuationSizes : null);
                    break;

                case BiffRecordType.Note:
                    _comments?.OnNote(body);
                    break;

                case BiffRecordType.WsBool:
                    _pageSetup.OnWsBool(body);
                    break;

                case BiffRecordType.Setup:
                    _pageSetup.OnSetup(body);
                    break;

                case BiffRecordType.LeftMargin or BiffRecordType.RightMargin
                    or BiffRecordType.TopMargin or BiffRecordType.BottomMargin:
                    _pageSetup.OnMargin(reader.RecordType, body);
                    break;

                case BiffRecordType.Header or BiffRecordType.Footer:
                    _pageSetup.OnHeaderFooterText(reader.RecordType, body);
                    break;

                case BiffRecordType.HeaderFooter:
                    _pageSetup.OnHeaderFooterExtension(body);
                    break;

                case BiffRecordType.HCenter or BiffRecordType.VCenter
                    or BiffRecordType.PrintGrid or BiffRecordType.PrintRowCol:
                    _pageSetup.OnPrintFlag(reader.RecordType, body);
                    break;

                case BiffRecordType.HorizontalPageBreaks or BiffRecordType.VerticalPageBreaks:
                    _pageSetup.OnPageBreaks(reader.RecordType, body);
                    break;

                case BiffRecordType.CondFmt or BiffRecordType.CondFmt12:
                    _conditionalFormats.OnCondFmt(body, reader.RecordType == BiffRecordType.CondFmt12);
                    break;

                case BiffRecordType.Cf:
                    _conditionalFormats.OnCf(body, decoder, sheetIndex);
                    break;

                case BiffRecordType.Cf12:
                    _conditionalFormats.OnCf12(body, decoder, sheetIndex);
                    break;

                case BiffRecordType.CfEx:
                    _conditionalFormats.OnCfEx(body);
                    break;

                case BiffRecordType.AutoFilterInfo:
                    // The range itself lives in the _FilterDatabase name, applied by the loader.
                    break;

                default:
                    break;
            }
        }

        Finish(sink);
        if (hasFrozenPane)
            sink.OnFrozenPane(frozenRows, frozenColumns);
    }

    /// <summary>
    /// Reads past a substream nested in this one, leaving the reader on its end marker. The
    /// nesting is counted because a chart may hold substreams of its own.
    /// </summary>
    private void SkipNestedSubstream()
    {
        int depth = 1;
        while (depth > 0 && reader.RecordOffset < substreamEnd && reader.Read())
        {
            if (reader.RecordType == BiffRecordType.Bof)
                depth++;
            else if (reader.RecordType == BiffRecordType.Eof)
                depth--;
        }
    }

    /// <summary>Applies everything that could only be resolved once the substream ended.</summary>
    private void Finish(WorksheetContentSink sink)
    {
        FinishFormulas(sink);
        _comments?.Finish(sink.Sheet, warn);
        _pageSetup.Finish(sink.Sheet);
        _conditionalFormats.Finish(sink.Sheet, warn);
        FinishPictures(sink.Sheet);
    }

    /// <summary>
    /// Joins the sheet's drawing records and takes the pictures out of them. Nothing happens
    /// when the workbook had no drawing group: without the store there are no bytes to attach,
    /// only shapes naming images that are not there.
    /// </summary>
    private void FinishPictures(Worksheet sheet)
    {
        if (_drawing.Count == 0 || globals.Blips.Count == 0)
            return;

        var shapes = new List<EscherPictures.Shape>();
        EscherPictures.ReadShapes(EscherPictures.Join(_drawing), shapes);
        EscherPictures.Attach(sheet, globals.Blips, shapes);
        _drawing.Clear();
    }

    private void ReadRow(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 16)
            return;

        int row = BinaryPrimitives.ReadUInt16LittleEndian(body) + 1;
        ushort heightField = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body[12..]);

        // The top bit of the height says the row keeps the sheet default rather than its own.
        bool custom = (flags & 0x00000040) != 0;
        double? height = custom || (heightField & 0x8000) == 0 ? (heightField & 0x7FFF) / 20.0 : null;

        bool hasStyle = (flags & 0x00000080) != 0;
        int styleIndex = hasStyle ? (int)((flags >> 16) & 0x0FFF) : 0;

        var definition = new RowDefinition(
            Row: row,
            Height: height,
            Hidden: (flags & 0x00000020) != 0,
            OutlineLevel: (int)(flags & 0x00000007),
            Collapsed: (flags & 0x00000010) != 0,
            RawStyleIndex: 0,
            ThickTop: (flags & 0x10000000) != 0,
            ThickBottom: (flags & 0x20000000) != 0);

        if (definition.IsEmpty && !hasStyle)
            return;

        sink.OnRow(definition);
        if (hasStyle)
            sink.Sheet.RowTable[row] = sink.Sheet.RowTable[row] with { Style = globals.StyleOf(styleIndex) };
    }

    private void ReadBlank(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 6)
            return;

        SetCell(sink, ReadCellRef(body), CellValue.Blank, BinaryPrimitives.ReadUInt16LittleEndian(body[4..]));
    }

    private void ReadMulBlank(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 6)
            return;

        int row = BinaryPrimitives.ReadUInt16LittleEndian(body) + 1;
        int firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        int count = (body.Length - 6) / 2;
        for (int i = 0; i < count; i++)
        {
            int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[(4 + (i * 2))..]);
            SetCell(sink, new CellRef(row, firstColumn + i), CellValue.Blank, xf);
        }
    }

    private void ReadRk(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 10)
            return;

        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        double number = BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(body[6..]));
        SetCell(sink, ReadCellRef(body), MakeNumeric(number, xf), xf);
    }

    private void ReadMulRk(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 6)
            return;

        int row = BinaryPrimitives.ReadUInt16LittleEndian(body) + 1;
        int firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        int count = (body.Length - 6) / 6;
        for (int i = 0; i < count; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(4 + (i * 6), 6);
            int xf = BinaryPrimitives.ReadUInt16LittleEndian(entry);
            double number = BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(entry[2..]));
            SetCell(sink, new CellRef(row, firstColumn + i), MakeNumeric(number, xf), xf);
        }
    }

    private void ReadNumber(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 14)
            return;

        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        double number = BinaryPrimitives.ReadDoubleLittleEndian(body[6..]);
        SetCell(sink, ReadCellRef(body), MakeNumeric(number, xf), xf);
    }

    private void ReadBoolErr(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return;

        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        CellValue value = body[7] != 0
            ? CellValue.Error(BiffValues.DecodeError(body[6]))
            : CellValue.Boolean(body[6] != 0);
        SetCell(sink, ReadCellRef(body), value, xf);
    }

    private void ReadLabelSst(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 10)
            return;

        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int index = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[6..]);
        CellRef cell = ReadCellRef(body);
        CellValue textValue = MakeText(globals.StringAt(index), cell, sink, accountText: false);
        SetCell(sink, cell, textValue, xf);
        if (globals.RichStringAt(index) is { IsFormatted: true } rich &&
            textValue.GetText().Length == rich.Text.Length)
        {
            sink.Sheet.RichStringTable[cell] = new IO.RichStringMarkup(rich.Text, Streaming.RichTextMarkup.ToMarkup(rich));
        }
    }

    /// <summary>
    /// Wraps text as a value, cutting it to what a cell can hold. Excel itself enforces the
    /// limit, but files exist that exceed it, and refusing to load one over a single long cell
    /// would be the wrong trade.
    /// </summary>
    private CellValue MakeText(
        string text, CellRef cell, WorksheetContentSink sink, bool accountText = true)
    {
        if (accountText)
            loadBudget?.AddCellText(text);

        if (text.Length <= SheetLimits.MaxCellTextLength)
            return CellValue.Text(text);

        warn(new WorkbookWarning(
            "TextTooLong",
            $"A cell of '{sink.Sheet.Name}' holds {text.Length} characters, past the {SheetLimits.MaxCellTextLength} " +
            "a cell may have; the text was cut short.",
            cell));
        return CellValue.Text(text[..SheetLimits.MaxCellTextLength]);
    }

    private void ReadLabel(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return;

        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int cursor = 6;
        string text = globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadUnicodeString(body, ref cursor, loadBudget)
            : BiffStrings.ReadByteString(body, ref cursor, globals.Encoding, lengthBytes: 2);
        CellRef cell = ReadCellRef(body);
        CellValue textValue = MakeText(text, cell, sink);
        SetCell(sink, cell, textValue, xf);

        if (reader.RecordType == BiffRecordType.RString &&
            globals.Version >= BiffVersion.Biff8 &&
            cursor + 2 <= body.Length)
        {
            int runCount = BinaryPrimitives.ReadUInt16LittleEndian(body[cursor..]);
            cursor += 2;
            int available = Math.Min(runCount, (body.Length - cursor) / 4);
            if (available > 0)
            {
                var runs = new (int CharPos, int FontIndex)[available];
                for (int i = 0; i < available; i++)
                {
                    runs[i] = (
                        BinaryPrimitives.ReadUInt16LittleEndian(body[cursor..]),
                        BinaryPrimitives.ReadUInt16LittleEndian(body[(cursor + 2)..]));
                    cursor += 4;
                }

                if (globals.BuildRichText(text, runs) is { IsFormatted: true } rich &&
                    textValue.GetText().Length == rich.Text.Length)
                {
                    sink.Sheet.RichStringTable[cell] = new IO.RichStringMarkup(rich.Text, Streaming.RichTextMarkup.ToMarkup(rich));
                }
            }
        }
    }

    private void ReadFormula(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 22)
            return;

        CellRef cell = ReadCellRef(body);
        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        bool needsString = BiffValues.TryReadCachedValue(body[6..14], out CellValue value);

        // A formula's cached number is a date exactly as often as a literal one is, and the
        // format table is the only thing that says so here too.
        if (value.Kind == CellValueKind.Number)
            value = MakeNumeric(value.GetNumber(), xf);

        int rgceLength = BinaryPrimitives.ReadUInt16LittleEndian(body[20..]);
        rgceLength = Math.Min(rgceLength, Math.Max(0, body.Length - 22));
        ReadOnlySpan<byte> rgce = body.Slice(22, rgceLength);

        // A cell that only repeats a formula stated elsewhere points at the master with PtgExp;
        // the master may not have been read yet, so the link is resolved at the end.
        if (rgceLength >= 5 && rgce[0] == 0x01)
        {
            var master = new CellRef(
                BinaryPrimitives.ReadUInt16LittleEndian(rgce[1..]) + 1,
                BinaryPrimitives.ReadUInt16LittleEndian(rgce[3..]) + 1);
            _sharedMembers.Add((cell, master));
            SetCell(sink, cell, value, xf);
            if (needsString)
                MarkPendingString(cell);
            return;
        }

        string? formula = decoder.Decode(rgce, cell, sheetIndex);
        if (formula is not null)
            loadBudget?.AddCellText(formula);
        if (formula is null && rgceLength > 0)
        {
            warn(new WorkbookWarning(
                "UndecodedFormula",
                $"A formula in '{sink.Sheet.Name}' uses a construct this library cannot render; its last value was kept instead.",
                cell));
        }

        SetCell(sink, cell, value, xf, formula);
        if (needsString)
            MarkPendingString(cell);
    }

    private void MarkPendingString(CellRef cell)
    {
        _pendingStringCell = cell;
        _hasPendingString = true;
    }

    private void ReadPendingString(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (!_hasPendingString)
            return;

        int cursor = 0;
        string text = globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadUnicodeString(body, ref cursor, loadBudget)
            : BiffStrings.ReadByteString(body, ref cursor, globals.Encoding, lengthBytes: 2);
        sink.Sheet.Values.Set(
            _pendingStringCell.Row, _pendingStringCell.Column, MakeText(text, _pendingStringCell, sink));
        _hasPendingString = false;
    }

    private void ReadSharedFormula(ReadOnlySpan<byte> body)
    {
        if (body.Length < 10)
            return;

        var master = new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
            body[4] + 1);
        int rgceLength = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);
        rgceLength = Math.Min(rgceLength, Math.Max(0, body.Length - 10));
        _sharedFormulas[master] = body.Slice(10, rgceLength).ToArray();
    }

    private void ReadArrayFormula(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 14)
            return;

        var range = new RangeRef(
            new CellRef(BinaryPrimitives.ReadUInt16LittleEndian(body) + 1, body[4] + 1),
            new CellRef(BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1, body[5] + 1));
        int rgceLength = BinaryPrimitives.ReadUInt16LittleEndian(body[12..]);
        rgceLength = Math.Min(rgceLength, Math.Max(0, body.Length - 14));

        string? formula = decoder.Decode(body.Slice(14, rgceLength), range.Start, sheetIndex);
        if (formula is null)
            return;

        loadBudget?.AddCellText(formula);
        sink.OnFormulaGroup(new FormulaGroup(FormulaGroupKind.Array, range.Start, range, formula), sharedIndex: null);
        sink.Sheet.Formulas.Set(range.Start.Row, range.Start.Column, formula);
    }

    /// <summary>
    /// Resolves the cells that pointed at a shared master. The master's expression is written
    /// against its own position, so each member re-decodes it from where that member sits.
    /// </summary>
    private void FinishFormulas(WorksheetContentSink sink)
    {
        foreach ((CellRef cell, CellRef master) in _sharedMembers)
        {
            if (!_sharedFormulas.TryGetValue(master, out byte[]? rgce))
                continue;

            string? formula = decoder.Decode(rgce, cell, sheetIndex);
            if (formula is not null)
            {
                loadBudget?.AddCellText(formula);
                sink.Sheet.Formulas.Set(cell.Row, cell.Column, formula);
            }
        }

        foreach ((CellRef cell, byte[] rgce) in _deferredFormulas)
        {
            string? formula = decoder.Decode(rgce, cell, sheetIndex);
            if (formula is not null)
            {
                loadBudget?.AddCellText(formula);
                sink.Sheet.Formulas.Set(cell.Row, cell.Column, formula);
            }
        }

        _sharedMembers.Clear();
        _deferredFormulas.Clear();
    }

    private void ReadColumnInfo(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 10)
            return;

        int first = BinaryPrimitives.ReadUInt16LittleEndian(body) + 1;
        int last = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        int width = BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int xf = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);

        // A file may name every column up to the grid's edge just to set a default width; the
        // model stores those per column, so the run is clipped to what a sheet can hold.
        last = Math.Min(last, SheetLimits.MaxColumn);
        if (first > last)
            return;

        for (int column = first; column <= last; column++)
        {
            sink.OnColumn(new ColumnDefinition(
                Column: column,
                Width: width / 256.0,
                Hidden: (flags & 0x0001) != 0,
                OutlineLevel: (flags >> 8) & 0x07,
                Collapsed: (flags & 0x1000) != 0,
                RawStyleIndex: 0,
                BestFit: (flags & 0x0004) != 0));

            if (xf != 0 && xf != 15)
                sink.Sheet.ColumnTable[column] = sink.Sheet.ColumnTable[column] with { Style = globals.StyleOf(xf) };
        }
    }

    private void ReadMergeCells(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        int count = BinaryPrimitives.ReadUInt16LittleEndian(body);
        for (int i = 0; i < count && 2 + (i * 8) + 8 <= body.Length; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(2 + (i * 8), 8);
            var range = new RangeRef(
                new CellRef(
                    BinaryPrimitives.ReadUInt16LittleEndian(entry) + 1,
                    BinaryPrimitives.ReadUInt16LittleEndian(entry[4..]) + 1),
                new CellRef(
                    BinaryPrimitives.ReadUInt16LittleEndian(entry[2..]) + 1,
                    BinaryPrimitives.ReadUInt16LittleEndian(entry[6..]) + 1));
            sink.OnMerge(range);
        }
    }

    private static void ReadWindow2(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body);
        if ((flags & 0x0002) == 0)
            sink.OnGridLinesHidden();
    }

    private static void ReadZoom(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        if (body.Length < 4)
            return;

        int numerator = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int denominator = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        if (denominator == 0)
            return;

        int percent = numerator * 100 / denominator;
        if (percent is >= 10 and <= 400)
            sink.OnZoomScale(percent);
    }

    private static void ReadPane(
        ReadOnlySpan<byte> body, ref int frozenRows, ref int frozenColumns, ref bool hasFrozenPane)
    {
        if (body.Length < 4)
            return;

        frozenColumns = BinaryPrimitives.ReadUInt16LittleEndian(body);
        frozenRows = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        hasFrozenPane = frozenRows > 0 || frozenColumns > 0;
    }

    private static void ReadSheetExt(WorksheetContentSink sink, ReadOnlySpan<byte> body)
    {
        // The tab color sits at the end of the record, and only when the file bothered to set
        // one: a shorter record means the tab has the theme's default color.
        if (body.Length < 22)
            return;

        int colorIndex = (int)(BinaryPrimitives.ReadUInt32LittleEndian(body[16..]) & 0x7F);
        if (colorIndex is not (0 or 0x7F))
            sink.OnTabColor(ExcelColor.FromIndexed(colorIndex));
    }

    /// <summary>Reads the row and column that open every cell record.</summary>
    private static CellRef ReadCellRef(ReadOnlySpan<byte> body) => new(
        BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
        BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1);

    /// <summary>
    /// Wraps a number as the value its format says it is. A date in a spreadsheet is a number
    /// with a date format on it, and nothing else, so the format table is the only clue.
    /// </summary>
    private CellValue MakeNumeric(double number, int xfIndex) => globals.TemporalOf(xfIndex) switch
    {
        Styles.TemporalKind.Date => TemporalValue.AsDate(number, globals.DateSystem),
        Styles.TemporalKind.Duration => TemporalValue.AsDuration(number),
        _ => CellValue.Number(number),
    };

    private void SetCell(WorksheetContentSink sink, CellRef cell, in CellValue value, int xfIndex, string? formula = null)
    {
        if (!SheetLimits.IsValidRow(cell.Row) || !SheetLimits.IsValidColumn(cell.Column))
            return;

        sink.OnPooledCell(cell, value, globals.StyleOf(xfIndex), formula);
    }
}
