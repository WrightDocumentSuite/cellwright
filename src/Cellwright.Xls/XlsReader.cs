using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using Cellwright.Styles;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls;

/// <summary>
/// Reads the rows of a legacy workbook one at a time, without building a model of it. The
/// counterpart of <see cref="XlsxReader"/> for the binary format, and the way to walk a file
/// whose cells are wanted but whose formatting, formulas and drawings are not.
/// </summary>
/// <remarks>
/// The container the binary format lives in is not a forward-only thing: its stream is a chain
/// of sectors scattered through the file, and the sheet table at the front points at byte
/// offsets. So the workbook stream is held in memory whole, and what streaming buys is a
/// constant cost per row rather than a constant cost per file.
/// </remarks>
/// <example>
/// <code>
/// await using XlsReader reader = await XlsReader.OpenAsync("report.xls");
/// await foreach (RowData row in reader.ReadRowsAsync(reader.Sheets[0].Index))
///     Console.WriteLine(row.GetValueOrBlank(0));
/// </code>
/// </example>
public sealed class XlsReader : IAsyncDisposable
{
    private readonly byte[] _stream;
    private readonly IRecordDecryptor? _decryptor;
    private readonly XlsGlobals _globals;
    private readonly List<SheetInfo> _sheets = [];
    private readonly List<int> _sheetOffsets = [];
    private readonly RowBlock _block = new();
    private readonly DocumentLoadBudgetState _loadBudget;

    /// <summary>Where the formula sits whose text result arrives in the record after it.</summary>
    private int _pendingStringRow = -1;
    private int _pendingStringColumn = -1;
    private bool _disposed;

    private XlsReader(
        byte[] stream, IRecordDecryptor? decryptor, XlsGlobals globals, DocumentLoadBudgetState loadBudget)
    {
        _stream = stream;
        _decryptor = decryptor;
        _globals = globals;
        _loadBudget = loadBudget;

        foreach (XlsSheetEntry entry in globals.Sheets)
        {
            if (entry.Kind != XlsSheetKind.Worksheet)
                continue;

            _sheets.Add(new SheetInfo(_sheets.Count, entry.Name, entry.Visibility));
            _sheetOffsets.Add(entry.Offset);
        }
    }

    /// <summary>The worksheets of the workbook, in workbook order.</summary>
    public IReadOnlyList<SheetInfo> Sheets => _sheets;

    /// <summary>Date system the workbook counts its serial dates from.</summary>
    public ExcelDateSystem DateSystem => _globals.DateSystem;

    /// <summary>Opens a workbook from a file.</summary>
    /// <param name="path">Path of the .xls file.</param>
    /// <param name="options">Reading settings, chiefly the password of an encrypted file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    public static async ValueTask<XlsReader> OpenAsync(
        string path, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        await using FileStream stream = new(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 16, useAsync: true);
        return await OpenAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Opens a workbook from a stream.</summary>
    /// <param name="stream">Stream positioned at the start of the file.</param>
    /// <param name="options">Reading settings, chiefly the password of an encrypted file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The stream is not a legacy workbook, or is unreadable.</exception>
    public static async ValueTask<XlsReader> OpenAsync(
        Stream stream, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        options ??= new XlsLoadOptions();

        DocumentLoadBudget budget = options.EffectiveBudget;
        byte[] bytes = await XlsWorkbook.ReadAllAsync(stream, options.LeaveOpen, budget, cancellationToken).ConfigureAwait(false);
        var loadBudget = new DocumentLoadBudgetState(budget);
        (byte[] workbookStream, IRecordDecryptor? decryptor, XlsGlobals globals) =
            XlsWorkbookReader.OpenForStreaming(bytes, options, loadBudget, cancellationToken);
        return new XlsReader(workbookStream, decryptor, globals, loadBudget);
    }

    /// <summary>Enumerates the rows of a sheet located by name (case-insensitive).</summary>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="cancellationToken">Cancels the enumeration.</param>
    public IAsyncEnumerable<RowData> ReadRowsAsync(string sheetName, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(sheetName);
        foreach (SheetInfo sheet in _sheets)
        {
            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return ReadRowsAsync(sheet.Index, cancellationToken);
        }

        throw new ArgumentException($"The workbook has no sheet named '{sheetName}'.", nameof(sheetName));
    }

    /// <summary>
    /// Enumerates the rows of a sheet by zero-based index. Rows arrive in file order and empty
    /// ones are skipped, so watch <see cref="RowData.RowNumber"/>. The yielded
    /// <see cref="RowData"/> is a view over a reused buffer — copy it to keep it, and read one
    /// sheet at a time, since that buffer is the reader's.
    /// </summary>
    /// <param name="sheetIndex">Zero-based index into <see cref="Sheets"/>.</param>
    /// <param name="cancellationToken">Cancels the enumeration.</param>
    public async IAsyncEnumerable<RowData> ReadRowsAsync(
        int sheetIndex, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfNegative(sheetIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(sheetIndex, _sheets.Count);

        // Nothing here awaits: the bytes are already in hand. The signature is asynchronous so
        // that a caller can treat both formats the same way.
        await Task.CompletedTask.ConfigureAwait(false);

        using var reader = new BiffReader(_stream, _decryptor);
        reader.Seek(_sheetOffsets[sheetIndex]);

        // Where each row's last cell sits, so that a row can be handed over the moment nothing
        // more can be added to it. Nothing in the format stops a producer from coming back to
        // a row later, and files in the wild do, so the alternative to knowing is guessing.
        Dictionary<int, int> lastRecordOfRow = ScanRowEnds(
            reader, _sheetOffsets[sheetIndex], cancellationToken);

        _block.Reset();
        reader.Seek(_sheetOffsets[sheetIndex]);
        bool inSheet = false;

        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!inSheet)
            {
                if (reader.RecordType == BiffRecordType.Bof)
                    inSheet = true;
                continue;
            }

            if (reader.RecordType is BiffRecordType.Eof or BiffRecordType.Bof)
                break;

            // The record body is never held in a local: a span cannot survive the yield below.
            int offset = reader.RecordOffset;
            ReadRecord(reader.RecordType, reader.Record.Span);

            while (_block.TryTakeCompleted(lastRecordOfRow, offset, out RowData completed))
                yield return completed;
        }

        while (_block.TryTake(out RowData row))
            yield return row;
    }

    /// <summary>
    /// Walks a sheet's records once without decoding them, noting where each row's last cell
    /// record is. Only headers and row numbers are read, so the pass costs a fraction of the
    /// one that follows it.
    /// </summary>
    private static Dictionary<int, int> ScanRowEnds(
        BiffReader reader, int sheetOffset, CancellationToken cancellationToken)
    {
        var ends = new Dictionary<int, int>();
        reader.Seek(sheetOffset);

        bool inSheet = false;
        int lastRow = -1;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!inSheet)
            {
                if (reader.RecordType == BiffRecordType.Bof)
                    inSheet = true;
                continue;
            }

            if (reader.RecordType is BiffRecordType.Eof or BiffRecordType.Bof)
                break;

            // A formula's text result trails it in a record of its own, which belongs to the
            // same row even though it does not say so.
            if (reader.RecordType == BiffRecordType.String)
            {
                if (lastRow >= 0)
                    ends[lastRow] = reader.RecordOffset;
                continue;
            }

            int row = RowOf(reader.RecordType, reader.Record.Span);
            if (row < 0)
                continue;

            lastRow = row;
            ends[row] = reader.RecordOffset;
        }

        return ends;
    }

    /// <summary>Adds one record's cells to the block being gathered.</summary>
    private void ReadRecord(ushort recordType, ReadOnlySpan<byte> body)
    {
        // A formula whose result is text keeps that text in the record that follows it.
        if (recordType == BiffRecordType.String)
        {
            ReadPendingString(body);
            return;
        }

        int row = RowOf(recordType, body);
        if (row >= 0)
            ReadCells(recordType, body, row);
    }

    /// <summary>The one-based row a cell record belongs to, or -1 when it is not a cell.</summary>
    private static int RowOf(ushort recordType, ReadOnlySpan<byte> body) => recordType switch
    {
        BiffRecordType.Blank or BiffRecordType.MulBlank or BiffRecordType.Rk or BiffRecordType.MulRk or
        BiffRecordType.Number or BiffRecordType.BoolErr or BiffRecordType.LabelSst or BiffRecordType.Label or
        BiffRecordType.RString or BiffRecordType.Formula when body.Length >= 4 =>
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
        _ => -1,
    };

    /// <summary>Puts the values of one record into the block being gathered.</summary>
    private void ReadCells(ushort recordType, ReadOnlySpan<byte> body, int row)
    {
        switch (recordType)
        {
            case BiffRecordType.Rk when body.Length >= 10:
                Place(row, body, Numeric(BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(body[6..])), Xf(body)));
                break;

            case BiffRecordType.Number when body.Length >= 14:
                Place(row, body, Numeric(BinaryPrimitives.ReadDoubleLittleEndian(body[6..]), Xf(body)));
                break;

            case BiffRecordType.BoolErr when body.Length >= 8:
                Place(row, body, body[7] != 0
                    ? CellValue.Error(BiffValues.DecodeError(body[6]))
                    : CellValue.Boolean(body[6] != 0));
                break;

            case BiffRecordType.LabelSst when body.Length >= 10:
                Place(row, body, CellValue.Text(_globals.StringAt((int)BinaryPrimitives.ReadUInt32LittleEndian(body[6..]))));
                break;

            case BiffRecordType.Label or BiffRecordType.RString when body.Length >= 8:
                Place(row, body, ReadLabel(body));
                break;

            case BiffRecordType.Formula when body.Length >= 22:
                ReadFormulaValue(row, body);
                break;

            case BiffRecordType.Blank when body.Length >= 6:
                Place(row, body, CellValue.Blank);
                break;

            case BiffRecordType.MulRk when body.Length >= 6:
                ReadMultipleNumbers(row, body);
                break;

            case BiffRecordType.MulBlank when body.Length >= 6:
                ReadMultipleBlanks(row, body);
                break;

            default:
                break;
        }
    }

    private CellValue ReadLabel(ReadOnlySpan<byte> body)
    {
        int cursor = 6;
        string text = _globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadUnicodeString(body, ref cursor, _loadBudget)
            : BiffStrings.ReadByteString(body, ref cursor, _globals.Encoding, lengthBytes: 2);
        AccountText(text);
        return CellValue.Text(text);
    }

    /// <summary>
    /// Places a formula's last computed value. The expression is not decoded: a caller after
    /// rows wants what the cell shows, and that is the cached result.
    /// </summary>
    private void ReadFormulaValue(int row, ReadOnlySpan<byte> body)
    {
        bool textFollows = BiffValues.TryReadCachedValue(body[6..14], out CellValue value);
        if (value.Kind == CellValueKind.Number)
            value = Numeric(value.GetNumber(), Xf(body));

        Place(row, body, value);
        _pendingStringRow = textFollows ? row : -1;
        _pendingStringColumn = textFollows ? BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1 : -1;
    }

    /// <summary>Fills in the text result of the formula that came just before.</summary>
    private void ReadPendingString(ReadOnlySpan<byte> body)
    {
        if (_pendingStringColumn < 0)
            return;

        int cursor = 0;
        string text = _globals.Version >= BiffVersion.Biff8
            ? BiffStrings.ReadUnicodeString(body, ref cursor, _loadBudget)
            : BiffStrings.ReadByteString(body, ref cursor, _globals.Encoding, lengthBytes: 2);
        AccountText(text);
        _block.Place(_pendingStringRow, _pendingStringColumn, CellValue.Text(text));
        _pendingStringColumn = -1;
    }

    /// <summary>Reads a run of numbers packed into one record, each with its own format.</summary>
    private void ReadMultipleNumbers(int row, ReadOnlySpan<byte> body)
    {
        int firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        int count = (body.Length - 6) / 6;
        for (int i = 0; i < count; i++)
        {
            ReadOnlySpan<byte> entry = body.Slice(4 + (i * 6), 6);
            double number = BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(entry[2..]));
            Place(row, firstColumn + i, Numeric(number, BinaryPrimitives.ReadUInt16LittleEndian(entry)));
        }
    }

    /// <summary>Reads a run of empty but formatted cells, which carry no value at all.</summary>
    private void ReadMultipleBlanks(int row, ReadOnlySpan<byte> body)
    {
        int firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1;
        int count = (body.Length - 6) / 2;
        for (int i = 0; i < count; i++)
            Place(row, firstColumn + i, CellValue.Blank);
    }

    private static int Xf(ReadOnlySpan<byte> body) => BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);

    /// <summary>Wraps a number as the value its format says it is: a date, a duration or a number.</summary>
    private CellValue Numeric(double number, int xfIndex) => _globals.TemporalOf(xfIndex) switch
    {
        TemporalKind.Date => TemporalValue.AsDate(number, _globals.DateSystem),
        TemporalKind.Duration => TemporalValue.AsDuration(number),
        _ => CellValue.Number(number),
    };

    private void Place(int row, ReadOnlySpan<byte> body, in CellValue value) =>
        Place(row, BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1, value);

    private void Place(int row, int column, in CellValue value)
    {
        if (row < 1 || column < 1 || column > SheetLimits.MaxColumn)
            return;

        _loadBudget.AddCell();
        _block.Place(row, column, value);
    }

    private void AccountText(string text)
    {
        long item = 0;
        _loadBudget.AddCellText(text.Length, ref item);
    }

    /// <inheritdoc/>
    public ValueTask DisposeAsync()
    {
        _disposed = true;
        return ValueTask.CompletedTask;
    }

    /// <summary>
    /// Gathers the cells of a row block until it is complete. Within a block a producer may
    /// write cells in whatever order suits it — files in the wild go column-run by column-run
    /// across two rows at a time — so nothing can be handed over until the block's index
    /// record says the block is done.
    /// </summary>
    private sealed class RowBlock
    {
        private readonly Dictionary<int, Buffered> _rows = [];

        /// <summary>The buffered rows by number, so that the lowest is always to hand.</summary>
        private readonly PriorityQueue<int, int> _order = new();
        private readonly Stack<CellValue[]> _spare = new();
        private Buffered? _inFlight;

        /// <summary>One row's cells and how far along it they reach.</summary>
        private sealed class Buffered
        {
            public required CellValue[] Cells { get; set; }

            public int Width { get; set; }
        }

        /// <summary>Drops everything, ready for another sheet.</summary>
        public void Reset()
        {
            foreach (Buffered row in _rows.Values)
                Recycle(row);

            _rows.Clear();
            _order.Clear();
            if (_inFlight is not null)
            {
                Recycle(_inFlight);
                _inFlight = null;
            }
        }

        /// <summary>Records one cell.</summary>
        public void Place(int row, int column, in CellValue value)
        {
            if (!_rows.TryGetValue(row, out Buffered? buffered))
            {
                buffered = new Buffered { Cells = _spare.Count > 0 ? _spare.Pop() : new CellValue[64] };
                _rows[row] = buffered;
                _order.Enqueue(row, row);
            }

            if (column > buffered.Cells.Length)
            {
                CellValue[] grown = buffered.Cells;
                Array.Resize(ref grown, Math.Max(column, buffered.Cells.Length * 2));
                buffered.Cells = grown;
            }

            buffered.Cells[column - 1] = value;
            buffered.Width = Math.Max(buffered.Width, column);
        }

        /// <summary>
        /// Takes the lowest row if the record just read was its last. Rows are handed over in
        /// order, so a row that is finished waits for the ones before it.
        /// </summary>
        /// <param name="lastRecordOfRow">Where each row's final cell record sits.</param>
        /// <param name="offset">Offset of the record just read.</param>
        /// <param name="row">The row, when one was ready.</param>
        public bool TryTakeCompleted(Dictionary<int, int> lastRecordOfRow, int offset, out RowData row)
        {
            row = default;
            return _order.TryPeek(out int number, out _) &&
                lastRecordOfRow.TryGetValue(number, out int last) &&
                last <= offset &&
                TryTake(out row);
        }

        /// <summary>
        /// Takes the lowest row gathered so far. The buffer behind the row handed over last is
        /// reclaimed here, which is the contract <see cref="RowData"/> states: a row is a view
        /// that the next step of the enumeration invalidates.
        /// </summary>
        public bool TryTake(out RowData row)
        {
            row = default;
            if (!_order.TryDequeue(out int number, out _))
                return false;

            Buffered buffered = _rows[number];
            _rows.Remove(number);
            if (_inFlight is not null)
                Recycle(_inFlight);

            _inFlight = buffered;
            row = new RowData(buffered.Cells, buffered.Width, number);
            return true;
        }

        private void Recycle(Buffered row)
        {
            Array.Clear(row.Cells, 0, row.Width);
            row.Width = 0;
            _spare.Push(row.Cells);
        }
    }
}
