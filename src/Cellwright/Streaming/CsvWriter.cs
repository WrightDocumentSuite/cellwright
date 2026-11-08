using System.Globalization;
using System.Text;
using Cellwright.Styles;

namespace Cellwright;

/// <summary>
/// Streaming RFC 4180 CSV and TSV writer.
/// </summary>
public sealed class CsvWriter : IAsyncDisposable, IDisposable
{
    private readonly TextWriter _writer;
    private readonly CsvOptions _options;
    private readonly bool _ownsWriter;
    private bool _disposed;

    private CsvWriter(TextWriter writer, CsvOptions options, bool ownsWriter)
    {
        _writer = writer;
        _options = options;
        _ownsWriter = ownsWriter;
    }

    /// <summary>Creates a <see cref="CsvWriter"/> over a file path.</summary>
    public static CsvWriter Create(string path, CsvOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        CsvOptions effective = (options ?? CsvOptions.Default) with { LeaveOpen = false };
        effective.Validate();
        var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 64 * 1024, useAsync: true);
        var writer = new StreamWriter(stream, effective.Encoding, bufferSize: 16 * 1024, leaveOpen: false);
        return new CsvWriter(writer, effective, ownsWriter: true);
    }

    /// <summary>Creates a <see cref="CsvWriter"/> over a stream.</summary>
    public static CsvWriter Create(Stream stream, CsvOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        CsvOptions effective = options ?? CsvOptions.Default;
        effective.Validate();
        var writer = new StreamWriter(stream, effective.Encoding, bufferSize: 16 * 1024, leaveOpen: effective.LeaveOpen);
        return new CsvWriter(writer, effective, ownsWriter: true);
    }

    /// <summary>Creates a <see cref="CsvWriter"/> over an existing <see cref="TextWriter"/>.</summary>
    public static CsvWriter Create(TextWriter writer, CsvOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(writer);
        CsvOptions effective = options ?? CsvOptions.Default;
        effective.Validate();
        return new CsvWriter(writer, effective, ownsWriter: !effective.LeaveOpen);
    }

    /// <summary>Writes one row of <see cref="CellValue"/> items synchronously.</summary>
    public void WriteRow(ReadOnlySpan<CellValue> values)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        for (int i = 0; i < values.Length; i++)
        {
            if (i > 0)
                _writer.Write(_options.Delimiter);
            WriteField(_writer, FormatCellValue(values[i], _options.Culture), _options);
        }

        _writer.Write(_options.NewLine);
    }

    /// <summary>Writes one row of <see cref="CellValue"/> items asynchronously.</summary>
    public async ValueTask WriteRowAsync(ReadOnlyMemory<CellValue> values, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var builder = new StringBuilder();
        ReadOnlySpan<CellValue> span = values.Span;
        for (int i = 0; i < span.Length; i++)
        {
            if (i > 0)
                builder.Append(_options.Delimiter);
            AppendField(builder, FormatCellValue(span[i], _options.Culture), _options);
        }

        builder.Append(_options.NewLine);
        await _writer.WriteAsync(builder.ToString().AsMemory(), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a worksheet's used range as CSV/TSV into a string.</summary>
    internal static string ExportWorksheet(Worksheet sheet, CsvOptions options)
    {
        options.Validate();
        if (sheet.UsedRange is not { } used)
            return string.Empty;

        var builder = new StringBuilder();
        for (int row = 1; row <= used.End.Row; row++)
        {
            for (int col = 1; col <= used.End.Column; col++)
            {
                if (col > 1)
                    builder.Append(options.Delimiter);

                Cell cell = sheet.Cell(row, col);
                string text = options.UseFormattedValues
                    ? NumberFormatRenderer.Format(cell.Value, cell.StyleValue.NumberFormat, options.Culture, sheet.Workbook.DateSystem)
                    : FormatCellValue(cell.Value, options.Culture);
                AppendField(builder, text, options);
            }

            builder.Append(options.NewLine);
        }

        return builder.ToString();
    }

    internal static string FormatCellValue(in CellValue value, CultureInfo culture) => value.Kind switch
    {
        CellValueKind.Blank => string.Empty,
        CellValueKind.Number => value.GetNumber().ToString(culture),
        CellValueKind.Text => value.GetText(),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.DateTime => FormatDateTime(value.GetDateTime(), culture),
        CellValueKind.TimeSpan => value.GetTimeSpan().ToString("c", culture),
        CellValueKind.Error => value.GetError().ToText(),
        _ => string.Empty,
    };

    private static string FormatDateTime(DateTime dt, CultureInfo culture) =>
        dt.TimeOfDay == TimeSpan.Zero
            ? dt.ToString("yyyy-MM-dd", culture)
            : dt.ToString("yyyy-MM-dd HH:mm:ss", culture);

    private static void WriteField(TextWriter writer, string value, CsvOptions options)
    {
        if (!NeedsQuotes(value, options))
        {
            writer.Write(value);
            return;
        }

        writer.Write(options.Quote);
        foreach (char c in value)
        {
            if (c == options.Quote)
                writer.Write(options.Quote);
            writer.Write(c);
        }

        writer.Write(options.Quote);
    }

    private static void AppendField(StringBuilder builder, string value, CsvOptions options)
    {
        if (!NeedsQuotes(value, options))
        {
            builder.Append(value);
            return;
        }

        builder.Append(options.Quote);
        foreach (char c in value)
        {
            if (c == options.Quote)
                builder.Append(options.Quote);
            builder.Append(c);
        }

        builder.Append(options.Quote);
    }

    private static bool NeedsQuotes(string value, CsvOptions options)
    {
        if (value.Length == 0)
            return false;
        if (value[0] is ' ' or '\t' || value[^1] is ' ' or '\t')
            return true;

        foreach (char c in value)
        {
            if (c == options.Delimiter || c == options.Quote || c is '\r' or '\n')
                return true;
        }

        return false;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _writer.Flush();
        if (_ownsWriter)
            _writer.Dispose();
    }

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;
        _disposed = true;
        await _writer.FlushAsync().ConfigureAwait(false);
        if (_ownsWriter)
            await _writer.DisposeAsync().ConfigureAwait(false);
    }
}
