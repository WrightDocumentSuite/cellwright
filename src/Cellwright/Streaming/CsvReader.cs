using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;

namespace Cellwright;

/// <summary>
/// Streaming RFC 4180 CSV and TSV reader yielding <see cref="RowData"/> rows with reusable cell
/// buffers, compatible with <c>[ExcelRow]</c> source-generated mappers.
/// </summary>
public static class CsvReader
{
    private static readonly string[] DateFormats =
        ["yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm"];

    private static readonly string[] DurationFormats =
        ["c", @"hh\:mm\:ss", @"h\:mm\:ss", @"\-hh\:mm\:ss", @"\-h\:mm\:ss"];

    /// <summary>Streams rows from a CSV/TSV file path.</summary>
    public static async IAsyncEnumerable<RowData> ReadRowsAsync(
        string path,
        CsvOptions? options = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        CsvOptions effective = (options ?? CsvOptions.Default) with { LeaveOpen = false };
        effective.Validate();

        var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 64 * 1024, useAsync: true);
        await using (stream.ConfigureAwait(false))
        {
            await foreach (RowData row in ReadRowsCoreAsync(stream, effective, cancellationToken).ConfigureAwait(false))
                yield return row;
        }
    }

    /// <summary>Streams rows from a CSV/TSV stream.</summary>
    public static IAsyncEnumerable<RowData> ReadRowsAsync(
        Stream stream,
        CsvOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        CsvOptions effective = options ?? CsvOptions.Default;
        effective.Validate();
        return ReadRowsCoreAsync(stream, effective, cancellationToken);
    }

    /// <summary>Reads rows synchronously from a <see cref="TextReader"/>.</summary>
    public static IEnumerable<RowData> ReadRows(TextReader reader, CsvOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(reader);
        CsvOptions effective = options ?? CsvOptions.Default;
        effective.Validate();
        return ReadRowsSyncCore(reader, effective);
    }

    private static async IAsyncEnumerable<RowData> ReadRowsCoreAsync(
        Stream stream,
        CsvOptions options,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (stream.CanSeek)
        {
            long remaining = Math.Max(0, stream.Length - stream.Position);
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxInputBytes), options.LoadBudget.MaxInputBytes, remaining);
        }

        // The limit is on bytes, and only the source stream knows how many there were: a decoded
        // character can be up to four of them. Seekable or not, the same input gets the same answer.
        using var reader = new StreamReader(
            new ByteLimitedStream(stream, options.LoadBudget.MaxInputBytes, options.LeaveOpen),
            options.Encoding,
            detectEncodingFromByteOrderMarks: true,
            bufferSize: 16 * 1024,
            leaveOpen: false);
        var parser = new StreamParser(options);

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            while (parser.TryReadRow(out RowData row))
                yield return row;

            int read = await reader.ReadAsync(parser.ChunkMemory, cancellationToken).ConfigureAwait(false);
            if (read == 0)
                break;
            parser.OnChunkRead(read);
        }

        if (parser.TryFlushFinalRow(out RowData finalRow))
            yield return finalRow;
    }

    private static IEnumerable<RowData> ReadRowsSyncCore(TextReader reader, CsvOptions options)
    {
        var parser = new StreamParser(options);
        while (true)
        {
            while (parser.TryReadRow(out RowData row))
                yield return row;

            int read = reader.Read(parser.ChunkArray, 0, parser.ChunkArray.Length);
            if (read == 0)
                break;
            parser.OnChunkRead(read);
        }

        if (parser.TryFlushFinalRow(out RowData finalRow))
            yield return finalRow;
    }

    internal static CellValue ConvertField(string raw, bool wasQuoted, CsvOptions options)
    {
        if (raw.Length == 0)
            return wasQuoted ? CellValue.Text(string.Empty) : CellValue.Blank;

        if (!options.InferTypes || (wasQuoted && !options.InferQuotedTypes))
            return CellValue.Text(raw);

        if (bool.TryParse(raw, out bool boolean))
            return CellValue.Boolean(boolean);

        if (ExcelErrorText.TryParse(raw, out ExcelError error))
            return CellValue.Error(error);

        if (TryParseNumberOrTemporal(raw, options.Culture, out CellValue parsed))
            return parsed;

        return CellValue.Text(raw);
    }

    private static bool TryParseNumberOrTemporal(string raw, CultureInfo culture, out CellValue value)
    {
        bool hasLeadingZeroDigits = raw.Length > 1 && raw[0] == '0' && char.IsAsciiDigit(raw[1]);
        if (!hasLeadingZeroDigits &&
            double.TryParse(raw, NumberStyles.Float, culture, out double number) &&
            double.IsFinite(number))
        {
            value = CellValue.Number(number);
            return true;
        }

        if (DateTime.TryParseExact(raw, DateFormats, culture, DateTimeStyles.None, out DateTime dateTime))
        {
            value = CellValue.DateTime(dateTime);
            return true;
        }

        if (raw.Contains(':', StringComparison.Ordinal) &&
            TimeSpan.TryParseExact(raw, DurationFormats, culture, out TimeSpan duration))
        {
            value = CellValue.TimeSpan(duration);
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>Counts the bytes read through it and stops at the input limit.</summary>
    private sealed class ByteLimitedStream(Stream inner, long maxBytes, bool leaveOpen) : Stream
    {
        private long _bytesRead;

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            Counted(inner.Read(buffer, offset, count));

        public override int Read(Span<byte> buffer) => Counted(inner.Read(buffer));

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Counted(await inner.ReadAsync(buffer.AsMemory(offset, count), cancellationToken).ConfigureAwait(false));

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            Counted(await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false));

        private int Counted(int read)
        {
            _bytesRead = DocumentLoadBudgetState.AddWithoutOverflow(_bytesRead, read);
            DocumentLoadBudgetState.Ensure(nameof(DocumentLoadBudget.MaxInputBytes), maxBytes, _bytesRead);
            return read;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing && !leaveOpen)
                inner.Dispose();

            base.Dispose(disposing);
        }
    }

    private sealed class StreamParser(CsvOptions options)
    {
        private readonly DocumentLoadBudgetState _budget = new(options.LoadBudget);
        private readonly StringBuilder _field = new(64);
        private CellValue[] _rowBuffer = new CellValue[64];
        private State _state = State.StartOfRecord;
        private int _chunkPos;
        private int _chunkLen;
        private int _col;
        private int _rowNumber;
        private bool _wasQuoted;
        private long _fieldChars;
        private long _totalInputChars;

        private enum State
        {
            StartOfRecord,
            StartOfField,
            UnquotedField,
            QuotedField,
            QuoteInQuotedField,
            TrailingAfterQuote,
            AfterCr,
        }

        public char[] ChunkArray { get; } = new char[8192];

        public Memory<char> ChunkMemory => ChunkArray;

        /// <summary>
        /// A <see cref="TextReader"/> has no bytes to count, so its characters are held to the
        /// input limit instead. For a stream this can never trip first: a character is at least a byte.
        /// </summary>
        public void OnChunkRead(int count)
        {
            _chunkPos = 0;
            _chunkLen = count;
            _totalInputChars = DocumentLoadBudgetState.AddWithoutOverflow(_totalInputChars, count);
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxInputBytes), _budget.Budget.MaxInputBytes, _totalInputChars);
        }

        public bool TryReadRow(out RowData row)
        {
            while (_chunkPos < _chunkLen)
            {
                char c = ChunkArray[_chunkPos++];
                if (StepChar(c, out row))
                    return true;
            }

            row = default;
            return false;
        }

        public bool TryFlushFinalRow(out RowData row)
        {
            if (_state is State.StartOfRecord or State.AfterCr)
            {
                row = default;
                return false;
            }

            if (options.StrictQuotes && _state == State.QuotedField)
                throw new InvalidDataException($"Row {_rowNumber + 1} ends inside a quoted field.");

            CompleteField();
            _state = State.StartOfRecord;
            return CompleteRecord(out row);
        }

        private bool StepChar(char c, out RowData row)
        {
            if (_state == State.AfterCr)
            {
                _state = State.StartOfRecord;
                if (c == '\n')
                {
                    row = default;
                    return false;
                }
            }

            return _state switch
            {
                State.StartOfRecord => StepStart(c, isRecordStart: true, out row),
                State.StartOfField => StepStart(c, isRecordStart: false, out row),
                State.UnquotedField => StepUnquoted(c, out row),
                State.QuotedField => StepQuoted(c, out row),
                State.QuoteInQuotedField => StepAfterQuote(c, out row),
                _ => StepTrailing(c, out row),
            };
        }

        private bool StepStart(char c, bool isRecordStart, out RowData row)
        {
            row = default;
            if (c == options.Quote)
            {
                _wasQuoted = true;
                _state = State.QuotedField;
                return false;
            }

            if (c == options.Delimiter)
            {
                CompleteField();
                _state = State.StartOfField;
                return false;
            }

            if (c is '\r' or '\n')
            {
                if (!isRecordStart)
                    CompleteField();
                _state = c == '\r' ? State.AfterCr : State.StartOfRecord;
                return CompleteRecord(out row);
            }

            AppendFieldChar(c);
            _state = State.UnquotedField;
            return false;
        }

        private bool StepUnquoted(char c, out RowData row)
        {
            if (c == options.Delimiter)
            {
                CompleteField();
                _state = State.StartOfField;
                row = default;
                return false;
            }

            if (c is '\r' or '\n')
            {
                CompleteField();
                _state = c == '\r' ? State.AfterCr : State.StartOfRecord;
                return CompleteRecord(out row);
            }

            AppendFieldChar(c);
            row = default;
            return false;
        }

        private bool StepQuoted(char c, out RowData row)
        {
            if (c == options.Quote)
                _state = State.QuoteInQuotedField;
            else
                AppendFieldChar(c);

            row = default;
            return false;
        }

        private bool StepAfterQuote(char c, out RowData row)
        {
            if (c == options.Quote)
            {
                AppendFieldChar(options.Quote);
                _state = State.QuotedField;
                row = default;
                return false;
            }

            return StepTrailing(c, out row);
        }

        private bool StepTrailing(char c, out RowData row)
        {
            if (c == options.Delimiter)
            {
                CompleteField();
                _state = State.StartOfField;
                row = default;
                return false;
            }

            if (c is '\r' or '\n')
            {
                CompleteField();
                _state = c == '\r' ? State.AfterCr : State.StartOfRecord;
                return CompleteRecord(out row);
            }

            if (options.StrictQuotes)
                throw new InvalidDataException($"Row {_rowNumber + 1} has text after the closing quote of a field.");

            // Dropping it would lose data without a trace; keeping it is what Excel does.
            AppendFieldChar(c);
            _state = State.TrailingAfterQuote;
            row = default;
            return false;
        }

        private void AppendFieldChar(char c)
        {
            _budget.AddCellText(1, ref _fieldChars);
            _field.Append(c);
        }

        private void CompleteField()
        {
            _budget.AddCell();
            if (_rowBuffer.Length < _col + 1)
                Array.Resize(ref _rowBuffer, Math.Max(_col + 1, _rowBuffer.Length * 2));

            _rowBuffer[_col++] = ConvertField(_field.ToString(), _wasQuoted, options);
            _field.Clear();
            _fieldChars = 0;
            _wasQuoted = false;
        }

        private bool CompleteRecord(out RowData row)
        {
            _rowNumber++;
            int count = _col;
            _col = 0;
            while (count > 0 && _rowBuffer[count - 1].IsBlank)
                count--;

            if (count > 0)
            {
                row = new RowData(_rowBuffer, count, _rowNumber);
                return true;
            }

            row = default;
            return false;
        }
    }
}
