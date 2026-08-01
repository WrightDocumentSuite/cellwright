using Cellwright.Xls.Cfb;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls;

/// <summary>
/// Reads and writes legacy Excel workbooks — the binary <c>.xls</c> format of Excel 97-2003 and
/// its predecessors.
/// </summary>
/// <remarks>
/// A workbook read here is an ordinary <see cref="Workbook"/>: edit it, save it as .xlsx, and
/// the conversion is done. <see cref="SaveAsync(Workbook, string, XlsSaveOptions?, CancellationToken)"/>
/// writes the binary format back for the cases where the old format is the requirement.
/// </remarks>
public static class XlsWorkbook
{
    /// <summary>Reads a legacy workbook from a file.</summary>
    /// <param name="path">Path of the .xls file.</param>
    /// <param name="options">Reading settings, chiefly the password of an encrypted file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The file is not a legacy workbook, or is unreadable.</exception>
    public static async ValueTask<Workbook> LoadAsync(
        string path, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        await using FileStream stream = new(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 16, useAsync: true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads a legacy workbook from a stream.</summary>
    /// <param name="stream">Stream positioned at the start of the file.</param>
    /// <param name="options">Reading settings, chiefly the password of an encrypted file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The stream is not a legacy workbook, or is unreadable.</exception>
    public static async ValueTask<Workbook> LoadAsync(
        Stream stream, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        options ??= new XlsLoadOptions();

        byte[] bytes = await ReadAllAsync(stream, options.LeaveOpen, options.Limits, cancellationToken).ConfigureAwait(false);
        return XlsWorkbookReader.Read(bytes, options);
    }

    /// <summary>Writes a workbook to a file in the Excel 97-2003 format.</summary>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="path">Path of the .xls file to create.</param>
    /// <param name="options">Writing settings.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static async ValueTask SaveAsync(
        Workbook workbook, string path, XlsSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentException.ThrowIfNullOrEmpty(path);

        await using FileStream stream = new(
            path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 1 << 16, useAsync: true);
        await SaveAsync(workbook, stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes a workbook to a stream in the Excel 97-2003 format.</summary>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="destination">Stream to write to.</param>
    /// <param name="options">Writing settings.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static async ValueTask SaveAsync(
        Workbook workbook, Stream destination, XlsSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(destination);

        byte[] content = Writing.XlsWorkbookWriter.Write(workbook, options ?? new XlsSaveOptions());

        // The document's properties are not BIFF records: they go in streams of their own
        // beside the workbook, and only exist in the file when there is something to say.
        var streams = new List<CompoundFileWriter.NamedStream> { new("Workbook", content) };
        if (Writing.PropertySetWriter.WriteSummary(workbook.Properties) is { } summary)
            streams.Add(new CompoundFileWriter.NamedStream(Writing.PropertySetWriter.SummaryStreamName, summary));
        if (Writing.PropertySetWriter.WriteDocumentSummary(workbook.Properties) is { } document)
        {
            streams.Add(new CompoundFileWriter.NamedStream(
                Writing.PropertySetWriter.DocumentSummaryStreamName, document));
        }

        await CompoundFileWriter.WriteAsync(destination, streams, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Whether the bytes at the start of a stream look like a legacy workbook.</summary>
    /// <param name="head">The first bytes of the file; eight are enough.</param>
    public static bool IsLegacyWorkbook(ReadOnlySpan<byte> head) => CompoundFile.HasSignature(head);

    /// <summary>
    /// Reads a stream into memory. The container is read back to front — the sheet table at
    /// the front points at byte offsets — so there is no reading it as it arrives.
    /// </summary>
    internal static async ValueTask<byte[]> ReadAllAsync(
        Stream stream, bool leaveOpen, PackageLimits limits, CancellationToken cancellationToken)
    {
        try
        {
            if (stream is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> buffer) &&
                buffer.Offset == 0 && buffer.Count == buffer.Array!.Length)
            {
                EnsureInputFits(buffer.Count, limits);
                return buffer.Array;
            }

            long remaining = stream.CanSeek ? Math.Max(0, stream.Length - stream.Position) : 0;
            if (stream.CanSeek)
                EnsureInputFits(remaining, limits);

            using var copy = stream.CanSeek && remaining <= int.MaxValue
                ? new MemoryStream((int)remaining)
                : new MemoryStream();
            byte[] chunk = System.Buffers.ArrayPool<byte>.Shared.Rent(81920);
            try
            {
                int read;
                while ((read = await stream.ReadAsync(chunk, cancellationToken).ConfigureAwait(false)) > 0)
                {
                    EnsureInputFits((long)copy.Length + read, limits);
                    await copy.WriteAsync(chunk.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                }
            }
            finally
            {
                System.Buffers.ArrayPool<byte>.Shared.Return(chunk);
            }

            return copy.ToArray();
        }
        finally
        {
            if (!leaveOpen)
                await stream.DisposeAsync().ConfigureAwait(false);
        }
    }

    private static void EnsureInputFits(long size, PackageLimits limits)
    {
        long limit = Math.Min(limits.MaxTotalSize, int.MaxValue);
        if (size > limit)
        {
            throw new XlsFormatException(
                $"The binary workbook is {size} bytes, past the {limit}-byte input limit. " +
                "Raise XlsLoadOptions.Limits if the file is genuinely this large.");
        }
    }
}
