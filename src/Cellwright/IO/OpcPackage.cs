using System.IO.Compression;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.IO;

/// <summary>
/// A focused Open Packaging Conventions wrapper over <see cref="ZipArchive"/> using the
/// .NET 10 asynchronous zip APIs. Read mode navigates parts and relationships; write mode
/// accumulates relationships and content types and emits them on dispose.
/// </summary>
internal sealed class OpcPackage : IAsyncDisposable
{
    private readonly ZipArchive _archive;
    private readonly bool _writeMode;
    private readonly Dictionary<string, ZipArchiveEntry> _entries;
    private readonly Dictionary<string, List<OpcRelationship>> _writeRelationships = new(OpcNames.Instance);
    private readonly HashSet<string> _writtenParts = new(OpcNames.Instance);
    private readonly ContentTypeMap _contentTypes;
    private bool _contentTypesParsed;
    private bool _disposed;
    private long _buffered;

    private OpcPackage(ZipArchive archive, bool writeMode)
    {
        _archive = archive;
        _writeMode = writeMode;
        _contentTypes = new ContentTypeMap();
        _entries = new Dictionary<string, ZipArchiveEntry>(OpcNames.Instance);
        if (!writeMode)
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
                _entries[OpcPath.ToPartPath(entry.FullName)] = entry;
        }
    }

    /// <summary>
    /// How much of one part, and of all of them together, may be held in memory. A zip says how
    /// large its entries expand to and is under no obligation to tell the truth, so a package
    /// built to be hostile turns a few megabytes on disk into as much memory as it likes. The
    /// defaults sit far above any real spreadsheet and well below what exhausts a server.
    /// </summary>
    public PackageLimits Limits { get; set; } = PackageLimits.Default;

    /// <summary>Opens an existing package for reading.</summary>
    public static async ValueTask<OpcPackage> OpenReadAsync(
        Stream stream, bool leaveOpen, CancellationToken cancellationToken, PackageLimits? limits = null)
    {
        ZipArchive archive = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen, entryNameEncoding: null, cancellationToken)
            .ConfigureAwait(false);
        return new OpcPackage(archive, writeMode: false) { Limits = limits ?? PackageLimits.Default };
    }

    /// <summary>Creates a new package for forward-only writing.</summary>
    public static async ValueTask<OpcPackage> CreateAsync(Stream stream, bool leaveOpen, CancellationToken cancellationToken)
    {
        ZipArchive archive = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Create, leaveOpen, entryNameEncoding: null, cancellationToken)
            .ConfigureAwait(false);
        var package = new OpcPackage(archive, writeMode: true);
        package._contentTypes.AddDefault("rels", XlsxSchema.ContentTypeRelationships);
        package._contentTypes.AddDefault("xml", "application/xml");
        return package;
    }

    /// <summary>Returns <see langword="true"/> when a part with the given name exists (read mode).</summary>
    public bool PartExists(string partPath) => _entries.ContainsKey(partPath);

    /// <summary>
    /// Whether a part has already gone into the package (write mode). A zip with the same entry
    /// twice is not a package any consumer will open, so anything copied in verbatim has to
    /// give way to what the writer generated.
    /// </summary>
    public bool PartWritten(string partPath) => _writtenParts.Contains(partPath);

    /// <summary>Every part name in the package, in archive order (read mode).</summary>
    public IEnumerable<string> PartPaths => _entries.Keys;

    /// <summary>Reads a whole part into memory (read mode).</summary>
    public async ValueTask<byte[]> ReadPartBytesAsync(string partPath, CancellationToken cancellationToken)
    {
        if (!_entries.TryGetValue(partPath, out ZipArchiveEntry? entry))
            throw new XlsxFormatException($"The package part '{partPath}' is missing.");

        using MemoryStream buffered = await BufferAsync(entry, cancellationToken).ConfigureAwait(false);
        return buffered.ToArray();
    }

    /// <summary>
    /// Opens a part for reading through a stream that refuses to expand past the per-part
    /// budget. Parts such as shared strings that remain represented in memory also contribute
    /// their actual decompressed bytes to the package-wide budget.
    /// </summary>
    /// <exception cref="XlsxFormatException">The part does not exist.</exception>
    public async ValueTask<Stream> OpenPartAsync(
        string partPath, CancellationToken cancellationToken, bool retainedInMemory = false)
    {
        if (!_entries.TryGetValue(partPath, out ZipArchiveEntry? entry))
            throw new XlsxFormatException($"The package part '{partPath}' is missing.");
        if (entry.Length > Limits.MaxPartSize)
            throw TooLarge(entry.FullName, entry.Length, Limits.MaxPartSize);

        Stream source = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        return new BoundedPartStream(
            source,
            entry.FullName,
            Limits.MaxPartSize,
            retainedInMemory ? CountBuffered : null);
    }

    /// <summary>Reads and resolves the relationships of a part. Returns an empty list when the part has none.</summary>
    public async ValueTask<IReadOnlyList<OpcRelationship>> GetRelationshipsAsync(string partPath, CancellationToken cancellationToken)
    {
        string relsPath = OpcPath.GetRelsPath(partPath);
        if (!_entries.TryGetValue(relsPath, out ZipArchiveEntry? entry))
            return [];

        using MemoryStream buffered = await BufferAsync(entry, cancellationToken).ConfigureAwait(false);
        return ParseRelationships(buffered);
    }

    /// <summary>Returns the content-type map of the package (read mode; parsed once).</summary>
    public async ValueTask<ContentTypeMap> GetContentTypesAsync(CancellationToken cancellationToken)
    {
        if (_contentTypesParsed)
            return _contentTypes;

        if (_entries.TryGetValue("/[Content_Types].xml", out ZipArchiveEntry? entry))
        {
            using MemoryStream buffered = await BufferAsync(entry, cancellationToken).ConfigureAwait(false);
            ContentTypeMap parsed = ContentTypeMap.Parse(buffered);
            _contentTypesParsed = true;
            return parsed;
        }

        _contentTypesParsed = true;
        return _contentTypes;
    }

    /// <summary>Creates a new part and registers its content type override (write mode).</summary>
    public async ValueTask<Stream> CreatePartAsync(string partPath, string contentType, CancellationToken cancellationToken)
    {
        RegisterContentType(partPath, contentType);
        _writtenParts.Add(partPath);
        ZipArchiveEntry entry = _archive.CreateEntry(OpcPath.ToEntryName(partPath), CompressionLevel.Optimal);
        return await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Copies a part into the package verbatim (write mode). The content type is expected to
    /// come from the preserved map or from an explicit <see cref="RegisterContentType"/> call.
    /// </summary>
    public async ValueTask WriteRawPartAsync(string partPath, byte[] content, CancellationToken cancellationToken)
    {
        _writtenParts.Add(partPath);
        ZipArchiveEntry entry = _archive.CreateEntry(OpcPath.ToEntryName(partPath), CompressionLevel.Optimal);
        Stream stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            await stream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Declares the content type of a part. Only needed when the extension default does not
    /// already cover it; <see cref="CreatePartAsync"/> does this on its own.
    /// </summary>
    public void RegisterContentType(string partPath, string contentType) =>
        _contentTypes.AddOverride(partPath, contentType);

    /// <summary>Seeds the content-type map with entries carried over from a loaded package (write mode).</summary>
    public void MergeContentTypes(ContentTypeMap source)
    {
        foreach ((string extension, string contentType) in source.Defaults)
            _contentTypes.AddDefault(extension, contentType);
        foreach ((string partPath, string contentType) in source.Overrides)
            _contentTypes.AddOverride(partPath, contentType);
    }

    /// <summary>Buffers a relationship to be written on dispose (write mode). Targets are written as given.</summary>
    public void AddRelationship(string sourcePartPath, OpcRelationship relationship)
    {
        if (!_writeRelationships.TryGetValue(sourcePartPath, out List<OpcRelationship>? list))
            _writeRelationships[sourcePartPath] = list = [];
        list.Add(relationship);
    }

    /// <summary>In write mode emits pending relationship parts and the content-types part, then closes the archive.</summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;
        _disposed = true;

        if (_writeMode)
        {
            foreach ((string sourcePart, List<OpcRelationship> relationships) in _writeRelationships)
                await WriteRelationshipsPartAsync(sourcePart, relationships).ConfigureAwait(false);

            // A content type for a part that never made it into the package makes the file
            // unopenable, so overrides carried over from a source package are pruned here.
            foreach (string partPath in _contentTypes.Overrides.Keys.ToArray())
            {
                if (!_writtenParts.Contains(partPath))
                    _contentTypes.RemoveOverride(partPath);
            }

            ZipArchiveEntry entry = _archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            Stream stream = await entry.OpenAsync(CancellationToken.None).ConfigureAwait(false);
            await using var writer = new Utf8XmlWriter(stream, bufferSize: 4 * 1024);
            _contentTypes.Write(writer);
        }

        await _archive.DisposeAsync().ConfigureAwait(false);
    }

    private async ValueTask WriteRelationshipsPartAsync(string sourcePart, List<OpcRelationship> relationships)
    {
        string relsPath = OpcPath.GetRelsPath(sourcePart);
        ZipArchiveEntry entry = _archive.CreateEntry(OpcPath.ToEntryName(relsPath), CompressionLevel.Optimal);
        Stream stream = await entry.OpenAsync(CancellationToken.None).ConfigureAwait(false);
        await using var writer = new Utf8XmlWriter(stream, bufferSize: 4 * 1024);

        writer.WriteDeclaration();
        writer.WriteRaw("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8);
        foreach (OpcRelationship relationship in relationships)
        {
            writer.WriteRaw("<Relationship Id=\""u8);
            writer.WriteAttributeText(relationship.Id);
            writer.WriteRaw("\" Type=\""u8);
            writer.WriteAttributeText(relationship.Type);
            writer.WriteRaw("\" Target=\""u8);
            writer.WriteAttributeText(relationship.Target);
            if (relationship.IsExternal)
                writer.WriteRaw("\" TargetMode=\"External"u8);
            writer.WriteRaw("\"/>"u8);
        }

        writer.WriteRaw("</Relationships>"u8);
    }

    /// <summary>
    /// Reads a part into memory, refusing one that will not fit the budget. The declared length
    /// only decides how large a buffer to start with; what actually arrives is what is counted.
    /// </summary>
    private async ValueTask<MemoryStream> BufferAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        long limit = Limits.MaxPartSize;
        if (entry.Length > limit)
            throw TooLarge(entry.FullName, entry.Length, limit);

        var buffered = new MemoryStream((int)Math.Clamp(entry.Length, 256, 1 << 20));
        Stream source = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using (source.ConfigureAwait(false))
        {
            await CopyBoundedAsync(source, buffered, entry.FullName, limit, cancellationToken).ConfigureAwait(false);
        }

        CountBuffered(buffered.Length);

        buffered.Position = 0;
        return buffered;
    }

    /// <summary>Copies a part, stopping the moment it grows past what it promised.</summary>
    private static async ValueTask CopyBoundedAsync(
        Stream source, Stream destination, string name, long limit, CancellationToken cancellationToken)
    {
        byte[] chunk = System.Buffers.ArrayPool<byte>.Shared.Rent(81920);
        try
        {
            long total = 0;
            int read;
            while ((read = await source.ReadAsync(chunk, cancellationToken).ConfigureAwait(false)) > 0)
            {
                total += read;
                if (total > limit)
                    throw TooLarge(name, total, limit);
                await destination.WriteAsync(chunk.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            }
        }
        finally
        {
            System.Buffers.ArrayPool<byte>.Shared.Return(chunk);
        }
    }

    private static XlsxFormatException TooLarge(string name, long size, long limit) => new(
        $"The package part '{name}' expands to at least {size} bytes, past the {limit}-byte limit. " +
        "Raise LoadOptions.Limits if the file is genuinely this large.");

    private void CountBuffered(long amount)
    {
        _buffered = checked(_buffered + amount);
        if (_buffered > Limits.MaxTotalSize)
        {
            throw new XlsxFormatException(
                $"The package expands past {Limits.MaxTotalSize} bytes in total, which no real workbook reaches. " +
                "Raise LoadOptions.Limits if the file is genuinely this large.");
        }
    }

    /// <summary>Counts what a zip actually returns rather than trusting its declared length.</summary>
    private sealed class BoundedPartStream(
        Stream source, string name, long limit, Action<long>? countRetained) : Stream
    {
        private long _read;

        public override bool CanRead => source.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            Account(source.Read(buffer, offset, count));

        public override int Read(Span<byte> buffer) => Account(source.Read(buffer));

        public override async ValueTask<int> ReadAsync(
            Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            Account(await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false));

        public override async Task<int> ReadAsync(
            byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Account(await source.ReadAsync(buffer.AsMemory(offset, count), cancellationToken).ConfigureAwait(false));

        public override int ReadByte()
        {
            int value = source.ReadByte();
            if (value >= 0)
                Account(1);
            return value;
        }

        private int Account(int count)
        {
            if (count <= 0)
                return count;

            _read = checked(_read + count);
            if (_read > limit)
                throw TooLarge(name, _read, limit);
            countRetained?.Invoke(count);
            return count;
        }

        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                source.Dispose();
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            await source.DisposeAsync().ConfigureAwait(false);
            GC.SuppressFinalize(this);
        }
    }

    private static List<OpcRelationship> ParseRelationships(Stream stream)
    {
        var relationships = new List<OpcRelationship>();
        using var reader = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;

            string? id = reader.GetAttribute("Id");
            string? type = reader.GetAttribute("Type");
            string? target = reader.GetAttribute("Target");
            if (id is null || type is null || target is null)
                continue;

            // A Strict package names the same roles under purl.oclc.org. Normalizing here means
            // everything above works in one vocabulary, and the package we write is transitional
            // throughout rather than half of each.
            bool isExternal = string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            relationships.Add(new OpcRelationship(id, XlsxSchema.Canonical(type), target, isExternal));
        }

        return relationships;
    }
}
