using System.Buffers;
using System.Buffers.Binary;
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
    private readonly DocumentLoadBudgetState? _loadBudget;
    // Zip permits duplicate names even though OPC does not. Track the physical entry object so
    // one duplicate cannot mark a later duplicate as already charged and bypass XML counters.
    private readonly HashSet<ZipArchiveEntry> _validatedXmlEntries = [];
    private readonly Dictionary<ZipArchiveEntry, long> _observedEntryBytes = [];
    private readonly HashSet<ZipArchiveEntry> _mediaEntries = [];
    private readonly Dictionary<ZipArchiveEntry, long> _chargedMediaBytes = [];
    private readonly HashSet<ZipArchiveEntry> _embeddedEntries = [];
    private ContentTypeMap? _parsedContentTypes;
    private bool _disposed;

    private OpcPackage(ZipArchive archive, bool writeMode, DocumentLoadBudgetState? loadBudget = null)
    {
        _archive = archive;
        _writeMode = writeMode;
        _loadBudget = loadBudget;
        _contentTypes = new ContentTypeMap();
        _entries = new Dictionary<string, ZipArchiveEntry>(OpcNames.Instance);
        if (!writeMode)
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
                _entries[OpcPath.ToPartPath(entry.FullName)] = entry;
        }
    }

    /// <summary>Opens an existing package for reading.</summary>
    public static async ValueTask<OpcPackage> OpenReadAsync(
        Stream stream,
        bool leaveOpen,
        CancellationToken cancellationToken,
        DocumentLoadBudget? budget = null)
    {
        var loadBudget = new DocumentLoadBudgetState(budget ?? DocumentLoadBudget.Default);
        return await OpenReadAsync(stream, leaveOpen, cancellationToken, loadBudget).ConfigureAwait(false);
    }

    /// <summary>Opens a package while continuing counters already used by an outer container.</summary>
    internal static async ValueTask<OpcPackage> OpenReadAsync(
        Stream stream,
        bool leaveOpen,
        CancellationToken cancellationToken,
        DocumentLoadBudgetState loadBudget)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(loadBudget);
        Stream packageStream;
        bool copied = !stream.CanSeek;
        if (copied)
        {
            try
            {
                byte[] content = await DocumentInput
                    .ReadBytesAsync(stream, loadBudget.Budget, cancellationToken)
                    .ConfigureAwait(false);
                packageStream = new MemoryStream(content, writable: false);
            }
            finally
            {
                if (!leaveOpen)
                    await stream.DisposeAsync().ConfigureAwait(false);
            }
        }
        else
        {
            long remaining = Math.Max(0, stream.Length - stream.Position);
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxInputBytes), loadBudget.Budget.MaxInputBytes, remaining);
            packageStream = stream;
        }

        bool archiveLeaveOpen = !copied && leaveOpen;
        ZipArchive archive;
        try
        {
            ValidateZipEntryCount(packageStream, loadBudget.Budget);
            archive = await ZipArchive.CreateAsync(
                    packageStream,
                    ZipArchiveMode.Read,
                    archiveLeaveOpen,
                    entryNameEncoding: null,
                    cancellationToken)
                .ConfigureAwait(false);
        }
        catch
        {
            if (copied || !leaveOpen)
                await packageStream.DisposeAsync().ConfigureAwait(false);
            throw;
        }

        try
        {
            var package = new OpcPackage(archive, writeMode: false, loadBudget);
            await package.ValidateEntriesAsync(cancellationToken).ConfigureAwait(false);
            return package;
        }
        catch
        {
            await archive.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    /// <summary>Checks central-directory metrics and semantic resource roles before allocation.</summary>
    private async ValueTask ValidateEntriesAsync(CancellationToken cancellationToken)
    {
        if (_loadBudget is null)
            return;

        DocumentLoadBudget budget = _loadBudget.Budget;
        _loadBudget.AddPackageParts(_archive.Entries.Count);

        long inflated = 0;
        foreach (ZipArchiveEntry entry in _archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxPartBytes), budget.MaxPartBytes, entry.Length);
            inflated = DocumentLoadBudgetState.AddWithoutOverflow(inflated, entry.Length);
        }
        _loadBudget.AddInflatedBytes(inflated);

        _parsedContentTypes = await TryReadContentTypesAsync(cancellationToken).ConfigureAwait(false);
        var aliases = new Dictionary<string, List<ZipArchiveEntry>>(OpcNames.Instance);
        foreach (ZipArchiveEntry entry in _archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string rawPath = OpcPath.ToPartPath(entry.FullName);
            string logicalPath = OpcPath.Unescape(rawPath);
            AddAlias(rawPath, entry);
            AddAlias(logicalPath, entry);

            string? contentType = _parsedContentTypes?.GetContentType(rawPath) ??
                                  _parsedContentTypes?.GetContentType(logicalPath);
            if (logicalPath.Contains("/media/", StringComparison.OrdinalIgnoreCase) ||
                IsMediaContentType(contentType))
            {
                _mediaEntries.Add(entry);
            }

            if (logicalPath.Contains("/embeddings/", StringComparison.OrdinalIgnoreCase))
                _embeddedEntries.Add(entry);
        }

        foreach (ZipArchiveEntry entry in _archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string relsPath = OpcPath.Unescape(OpcPath.ToPartPath(entry.FullName));
            if (!OpcPath.IsRelsPath(relsPath))
                continue;

            string sourcePart = OpcPath.GetSourcePart(relsPath);
            try
            {
                await ClassifyRelationshipsAsync(
                        entry,
                        sourcePart,
                        aliases,
                        cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (XmlException)
            {
                continue;
            }
        }

        foreach (ZipArchiveEntry entry in _mediaEntries)
            ChargeMedia(entry, ObservedSize(entry));
        foreach (ZipArchiveEntry entry in _embeddedEntries)
            _loadBudget.AddEmbeddedObject(ObservedSize(entry));

        void AddAlias(string alias, ZipArchiveEntry entry)
        {
            if (!aliases.TryGetValue(alias, out List<ZipArchiveEntry>? physicalEntries))
                aliases[alias] = physicalEntries = [];
            if (!physicalEntries.Contains(entry))
                physicalEntries.Add(entry);
        }
    }

    /// <summary>
    /// Reconciles an entry's actual decompressed size with the central-directory value already
    /// charged during preflight. Reopening the same part observes a maximum, not another copy.
    /// </summary>
    private void ObserveEntryBytes(ZipArchiveEntry entry, long observed)
    {
        if (_loadBudget is null)
            return;

        long previous = ObservedSize(entry);
        if (observed <= previous)
            return;

        _loadBudget.AddInflatedBytes(observed - previous);
        _observedEntryBytes[entry] = observed;
        if (_mediaEntries.Contains(entry))
            ChargeMedia(entry, observed);
        if (_embeddedEntries.Contains(entry))
        {
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxEmbeddedObjectBytes),
                _loadBudget.Budget.MaxEmbeddedObjectBytes,
                observed);
        }
    }

    private long ObservedSize(ZipArchiveEntry entry) =>
        _observedEntryBytes.TryGetValue(entry, out long observed)
            ? Math.Max(entry.Length, observed)
            : entry.Length;

    private void ChargeMedia(ZipArchiveEntry entry, long itemBytes)
    {
        if (_loadBudget is null)
            return;

        long previous = _chargedMediaBytes.GetValueOrDefault(entry);
        if (itemBytes <= previous)
            return;

        _loadBudget.AddMediaDelta(itemBytes, itemBytes - previous);
        _chargedMediaBytes[entry] = itemBytes;
    }

    private async ValueTask<ContentTypeMap?> TryReadContentTypesAsync(CancellationToken cancellationToken)
    {
        if (Locate("/[Content_Types].xml") is not { } entry)
            return null;

        var map = new ContentTypeMap();
        var physicalParts = new HashSet<string>(OpcNames.Instance);
        var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry candidate in _archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string rawPath = OpcPath.ToPartPath(candidate.FullName);
            physicalParts.Add(rawPath);
            physicalParts.Add(OpcPath.Unescape(rawPath));
            int dot = candidate.FullName.LastIndexOf('.');
            if (dot >= 0 && dot + 1 < candidate.FullName.Length)
                extensions.Add(candidate.FullName[(dot + 1)..]);
        }

        try
        {
            await ScanXmlEntryAsync(
                    entry,
                    cancellationToken,
                    reader =>
                    {
                        if (reader.NodeType != XmlNodeType.Element)
                            return;

                        if (reader.LocalName == "Default")
                        {
                            string? extension = reader.GetAttribute("Extension");
                            string? contentType = reader.GetAttribute("ContentType");
                            if (extension is not null && contentType is not null && extensions.Contains(extension))
                                map.AddDefault(extension, contentType);
                        }
                        else if (reader.LocalName == "Override")
                        {
                            string? partName = reader.GetAttribute("PartName");
                            string? contentType = reader.GetAttribute("ContentType");
                            if (partName is not null && contentType is not null &&
                                (physicalParts.Contains(partName) ||
                                 physicalParts.Contains(OpcPath.Unescape(partName))))
                            {
                                map.AddOverride(partName, contentType);
                            }
                        }
                    },
                    XmlSemantic.None)
                .ConfigureAwait(false);
            return map;
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private async ValueTask ClassifyRelationshipsAsync(
        ZipArchiveEntry entry,
        string sourcePart,
        IReadOnlyDictionary<string, List<ZipArchiveEntry>> aliases,
        CancellationToken cancellationToken)
    {
        await ScanXmlEntryAsync(
                entry,
                cancellationToken,
                reader =>
                {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                        return;

                    string? id = reader.GetAttribute("Id");
                    string? type = reader.GetAttribute("Type");
                    string? target = reader.GetAttribute("Target");
                    if (id is null || type is null || target is null)
                        return;

                    bool external = string.Equals(
                        reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                    if (external)
                        return;

                    MarkRelationshipTargets(
                        aliases,
                        sourcePart,
                        target,
                        IsMediaRelationship(type),
                        IsEmbeddedRelationship(type));
                },
                XmlSemantic.None)
            .ConfigureAwait(false);
    }

    private async ValueTask ValidateXmlAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        if (_loadBudget is null || _validatedXmlEntries.Contains(entry) || !LooksLikeXml(entry))
            return;

        try
        {
            await ScanXmlEntryAsync(
                    entry,
                    cancellationToken,
                    static _ => { },
                    SemanticOf(entry))
                .ConfigureAwait(false);
        }
        catch (XmlException)
        {
            // The format-specific reader retains its established malformed-XML diagnostics.
            // The valid prefix was still charged to the shared counters.
        }
    }

    private async ValueTask ScanXmlEntryAsync(
        ZipArchiveEntry entry,
        CancellationToken cancellationToken,
        Action<XmlReader> visit,
        XmlSemantic semantic)
    {
        if (_loadBudget is null || !_validatedXmlEntries.Add(entry))
            return;

        Stream raw = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        var bounded = new BoundedPartStream(
            raw,
            _loadBudget.Budget.MaxPartBytes,
            observed => ObserveEntryBytes(entry, observed));
        var content = new CancellationAwareReadStream(bounded, cancellationToken);
        await using (content.ConfigureAwait(false))
        {
            XmlReaderSettings settings = XmlDefaults.ForBudget(_loadBudget.Budget, async: true);
            using XmlReader reader = XmlReader.Create(content, settings);
            int itemDepth = -1;
            int textDepth = -1;
            string? cellType = null;
            long itemCharacters = 0;
            long formulaCharacters = 0;
            long valueCharacters = 0;
            bool formulaText = false;
            try
            {
                while (await reader.ReadAsync().ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    _loadBudget.AddXmlNode(reader.Depth + 1);

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (semantic == XmlSemantic.Worksheet && reader.LocalName == "c")
                        {
                            _loadBudget.AddCell();
                            itemDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            textDepth = -1;
                            cellType = reader.GetAttribute("t");
                            formulaCharacters = 0;
                            valueCharacters = 0;
                        }
                        else if (semantic == XmlSemantic.SharedStrings && reader.LocalName == "si")
                        {
                            _loadBudget.AddSharedString();
                            itemDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            itemCharacters = 0;
                        }
                        else if (semantic == XmlSemantic.Worksheet && itemDepth >= 0 &&
                                 (reader.LocalName == "f" ||
                                  (reader.LocalName == "t" && cellType == "inlineStr") ||
                                  (reader.LocalName == "v" && cellType == "str")))
                        {
                            textDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            formulaText = reader.LocalName == "f";
                        }
                    }
                    else if (itemDepth >= 0 && reader.Depth > itemDepth &&
                             (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace) &&
                             (semantic == XmlSemantic.SharedStrings ||
                              (textDepth >= 0 && reader.Depth > textDepth)))
                    {
                        if (semantic == XmlSemantic.SharedStrings)
                            _loadBudget.AddCellText(reader.Value.Length, ref itemCharacters);
                        else if (formulaText)
                            _loadBudget.AddCellText(reader.Value.Length, ref formulaCharacters);
                        else
                            _loadBudget.AddCellText(reader.Value.Length, ref valueCharacters);
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == textDepth)
                    {
                        textDepth = -1;
                        formulaText = false;
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == itemDepth)
                    {
                        itemDepth = -1;
                        cellType = null;
                    }

                    visit(reader);
                }
            }
            catch (XmlException exception) when (IsXmlCharacterLimit(exception))
            {
                long limit = _loadBudget.Budget.MaxXmlCharactersPerPart;
                throw new DocumentLoadLimitException(
                    nameof(DocumentLoadBudget.MaxXmlCharactersPerPart),
                    limit,
                    limit == long.MaxValue ? long.MaxValue : limit + 1);
            }
        }
    }

    private bool LooksLikeXml(ZipArchiveEntry entry)
    {
        string rawPath = OpcPath.ToPartPath(entry.FullName);
        string path = OpcPath.Unescape(rawPath);
        if (path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith(".vml", StringComparison.OrdinalIgnoreCase) ||
            path.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        string? contentType = _parsedContentTypes?.GetContentType(rawPath) ??
                              _parsedContentTypes?.GetContentType(path);
        return contentType is not null &&
               (contentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase) ||
                contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase));
    }

    private XmlSemantic SemanticOf(ZipArchiveEntry entry)
    {
        string rawPath = OpcPath.ToPartPath(entry.FullName);
        string path = OpcPath.Unescape(rawPath);
        string? contentType = _parsedContentTypes?.GetContentType(rawPath) ??
                              _parsedContentTypes?.GetContentType(path);
        if (string.Equals(contentType, XlsxSchema.ContentTypeWorksheet, StringComparison.OrdinalIgnoreCase) ||
            path.Contains("/worksheets/", StringComparison.OrdinalIgnoreCase))
        {
            return XmlSemantic.Worksheet;
        }

        return string.Equals(contentType, XlsxSchema.ContentTypeSharedStrings, StringComparison.OrdinalIgnoreCase) ||
               path.EndsWith("/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
            ? XmlSemantic.SharedStrings
            : XmlSemantic.None;
    }

    private void MarkRelationshipTargets(
        IReadOnlyDictionary<string, List<ZipArchiveEntry>> entries,
        string sourcePart,
        string target,
        bool media,
        bool embedded)
    {
        if (!media && !embedded)
            return;

        string resolved;
        try
        {
            resolved = OpcPath.Resolve(sourcePart, target);
        }
        catch (UriFormatException)
        {
            return;
        }

        Mark(resolved);
        string logical = OpcPath.Unescape(resolved);
        if (!OpcNames.Same(logical, resolved))
            Mark(logical);

        void Mark(string alias)
        {
            if (!entries.TryGetValue(alias, out List<ZipArchiveEntry>? physicalEntries))
                return;

            foreach (ZipArchiveEntry physicalEntry in physicalEntries)
            {
                if (media)
                    _mediaEntries.Add(physicalEntry);
                if (embedded)
                    _embeddedEntries.Add(physicalEntry);
            }
        }
    }

    private static bool IsMediaContentType(string? contentType) =>
        contentType is not null &&
        (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) ||
         contentType.StartsWith("audio/", StringComparison.OrdinalIgnoreCase) ||
         contentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase));

    private static bool IsMediaRelationship(string type)
    {
        string canonical = XlsxSchema.Canonical(type);
        return canonical == XlsxSchema.RelImage ||
               canonical.EndsWith("/audio", StringComparison.OrdinalIgnoreCase) ||
               canonical.EndsWith("/video", StringComparison.OrdinalIgnoreCase) ||
               canonical.EndsWith("/media", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEmbeddedRelationship(string type)
    {
        string canonical = XlsxSchema.Canonical(type);
        return canonical.EndsWith("/oleObject", StringComparison.OrdinalIgnoreCase) ||
               canonical.EndsWith("/package", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsXmlCharacterLimit(XmlException exception) =>
        exception.Message.Contains(nameof(XmlReaderSettings.MaxCharactersInDocument), StringComparison.Ordinal);

    private enum XmlSemantic
    {
        None,
        Worksheet,
        SharedStrings,
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
    public bool PartExists(string partPath) => Locate(partPath) is not null;

    /// <summary>Finds either the physical URI spelling or the decoded logical part name.</summary>
    private ZipArchiveEntry? Locate(string partPath)
    {
        if (_entries.TryGetValue(partPath, out ZipArchiveEntry? entry))
            return entry;

        string unescaped = OpcPath.Unescape(partPath);
        if (!unescaped.Equals(partPath, StringComparison.Ordinal) && _entries.TryGetValue(unescaped, out entry))
            return entry;

        foreach ((string candidate, ZipArchiveEntry value) in _entries)
        {
            if (OpcNames.Same(OpcPath.Unescape(candidate), unescaped))
                return value;
        }

        return null;
    }

    /// <summary>
    /// Whether a part has already gone into the package (write mode). A zip with the same entry
    /// twice is not a package any consumer will open, so anything copied in verbatim has to
    /// give way to what the writer generated.
    /// </summary>
    public bool PartWritten(string partPath) => _writtenParts.Contains(partPath);

    /// <summary>Every part name in the package, in archive order (read mode).</summary>
    public IEnumerable<string> PartPaths => _entries.Keys;

    /// <summary>Shared counters for format readers layered over this package.</summary>
    internal DocumentLoadBudgetState? LoadBudget => _loadBudget;

    /// <summary>Reads a whole part into memory (read mode).</summary>
    public async ValueTask<byte[]> ReadPartBytesAsync(string partPath, CancellationToken cancellationToken)
    {
        if (Locate(partPath) is not { } entry)
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
        _ = retainedInMemory; // Kept for source compatibility; every part is now charged at preflight.
        if (Locate(partPath) is not { } entry)
            throw new XlsxFormatException($"The package part '{partPath}' is missing.");

        DocumentLoadBudget budget = _loadBudget?.Budget ?? DocumentLoadBudget.Default;
        DocumentLoadBudgetState.Ensure(
            nameof(DocumentLoadBudget.MaxPartBytes), budget.MaxPartBytes, entry.Length);

        await ValidateXmlAsync(entry, cancellationToken).ConfigureAwait(false);

        Stream source = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        return new BoundedPartStream(
            source,
            budget.MaxPartBytes,
            observed => ObserveEntryBytes(entry, observed));
    }

    /// <summary>Reads and resolves the relationships of a part. Returns an empty list when the part has none.</summary>
    public async ValueTask<IReadOnlyList<OpcRelationship>> GetRelationshipsAsync(string partPath, CancellationToken cancellationToken)
    {
        string relsPath = OpcPath.GetRelsPath(partPath);
        if (Locate(relsPath) is not { } entry)
            return [];

        using MemoryStream buffered = await BufferAsync(entry, cancellationToken).ConfigureAwait(false);
        return ParseRelationships(buffered, cancellationToken);
    }

    /// <summary>Returns the content-type map of the package (read mode; parsed once).</summary>
    public async ValueTask<ContentTypeMap> GetContentTypesAsync(CancellationToken cancellationToken)
    {
        if (_parsedContentTypes is not null)
            return _parsedContentTypes;

        if (Locate("/[Content_Types].xml") is { } entry)
        {
            using MemoryStream buffered = await BufferAsync(entry, cancellationToken).ConfigureAwait(false);
            return _parsedContentTypes = ContentTypeMap.Parse(buffered, cancellationToken);
        }

        return _parsedContentTypes = _contentTypes;
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

    /// <summary>Rejects a ZIP entry-count bomb before ZipArchive materializes its directory.</summary>
    private static void ValidateZipEntryCount(Stream stream, DocumentLoadBudget budget)
    {
        if (!stream.CanSeek)
            return;

        long original = stream.Position;
        try
        {
            long available = Math.Max(0, stream.Length - original);
            int tailLength = (int)Math.Min(available, ushort.MaxValue + 22L);
            if (tailLength < 22)
                return;

            byte[] tail = ArrayPool<byte>.Shared.Rent(tailLength);
            try
            {
                long tailStart = stream.Length - tailLength;
                stream.Position = tailStart;
                stream.ReadExactly(tail.AsSpan(0, tailLength));

                int eocd = -1;
                for (int index = tailLength - 22; index >= 0; index--)
                {
                    if (BinaryPrimitives.ReadUInt32LittleEndian(tail.AsSpan(index, 4)) != 0x06054B50)
                        continue;

                    int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(tail.AsSpan(index + 20, 2));
                    if (index + 22 + commentLength <= tailLength)
                    {
                        eocd = index;
                        break;
                    }
                }

                if (eocd < 0)
                    return;

                ushort count = BinaryPrimitives.ReadUInt16LittleEndian(tail.AsSpan(eocd + 10, 2));
                if (count != ushort.MaxValue)
                {
                    DocumentLoadBudgetState.Ensure(
                        nameof(DocumentLoadBudget.MaxPackageParts), budget.MaxPackageParts, count);
                    return;
                }

                long eocdPosition = tailStart + eocd;
                if (eocdPosition - 20 < original)
                    return;

                Span<byte> locator = stackalloc byte[20];
                stream.Position = eocdPosition - locator.Length;
                stream.ReadExactly(locator);
                if (BinaryPrimitives.ReadUInt32LittleEndian(locator) != 0x07064B50)
                    return;

                ulong relativeOffset = BinaryPrimitives.ReadUInt64LittleEndian(locator[8..]);
                long maximumOffset = stream.Length - original - 56;
                if (maximumOffset < 0 || relativeOffset > (ulong)maximumOffset)
                    return;

                Span<byte> zip64 = stackalloc byte[56];
                stream.Position = original + (long)relativeOffset;
                stream.ReadExactly(zip64);
                if (BinaryPrimitives.ReadUInt32LittleEndian(zip64) != 0x06064B50)
                    return;

                ulong total = BinaryPrimitives.ReadUInt64LittleEndian(zip64[32..]);
                long observed = total > long.MaxValue ? long.MaxValue : (long)total;
                DocumentLoadBudgetState.Ensure(
                    nameof(DocumentLoadBudget.MaxPackageParts), budget.MaxPackageParts, observed);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(tail);
            }
        }
        finally
        {
            stream.Position = original;
        }
    }

    /// <summary>
    /// Reads a part into memory, refusing one that will not fit the budget. The declared length
    /// only decides how large a buffer to start with; what actually arrives is what is counted.
    /// </summary>
    private async ValueTask<MemoryStream> BufferAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        DocumentLoadBudget budget = _loadBudget?.Budget ?? DocumentLoadBudget.Default;
        long limit = budget.MaxPartBytes;
        DocumentLoadBudgetState.Ensure(nameof(DocumentLoadBudget.MaxPartBytes), limit, entry.Length);
        await ValidateXmlAsync(entry, cancellationToken).ConfigureAwait(false);

        var buffered = new MemoryStream((int)Math.Clamp(entry.Length, 256, 1 << 20));
        Stream source = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using (source.ConfigureAwait(false))
        {
            await CopyBoundedAsync(
                    source,
                    buffered,
                    limit,
                    observed => ObserveEntryBytes(entry, observed),
                    cancellationToken)
                .ConfigureAwait(false);
        }

        buffered.Position = 0;
        return buffered;
    }

    /// <summary>Copies a part, stopping the moment it grows past what it promised.</summary>
    private static async ValueTask CopyBoundedAsync(
        Stream source,
        Stream destination,
        long limit,
        Action<long> observe,
        CancellationToken cancellationToken)
    {
        byte[] chunk = System.Buffers.ArrayPool<byte>.Shared.Rent(81920);
        try
        {
            long total = 0;
            int read;
            while ((read = await source.ReadAsync(chunk, cancellationToken).ConfigureAwait(false)) > 0)
            {
                total = DocumentLoadBudgetState.AddWithoutOverflow(total, read);
                DocumentLoadBudgetState.Ensure(
                    nameof(DocumentLoadBudget.MaxPartBytes), limit, total);
                observe(total);
                await destination.WriteAsync(chunk.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            }
        }
        finally
        {
            System.Buffers.ArrayPool<byte>.Shared.Return(chunk);
        }
    }

    /// <summary>Counts what a zip actually returns rather than trusting its declared length.</summary>
    private sealed class BoundedPartStream(
        Stream source, long limit, Action<long>? observe) : Stream
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

            _read = DocumentLoadBudgetState.AddWithoutOverflow(_read, count);
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxPartBytes), limit, _read);
            observe?.Invoke(_read);
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

    /// <summary>Supplies the import token to XmlReader reads, including one giant text node.</summary>
    private sealed class CancellationAwareReadStream(Stream inner, CancellationToken cancellationToken) : Stream
    {
        public override bool CanRead => inner.CanRead;
        public override bool CanSeek => inner.CanSeek;
        public override bool CanWrite => false;
        public override long Length => inner.Length;
        public override long Position
        {
            get => inner.Position;
            set => inner.Position = value;
        }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = inner.Read(buffer, offset, count);
            cancellationToken.ThrowIfCancellationRequested();
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = inner.Read(buffer);
            cancellationToken.ThrowIfCancellationRequested();
            return read;
        }

        public override int ReadByte()
        {
            cancellationToken.ThrowIfCancellationRequested();
            int value = inner.ReadByte();
            cancellationToken.ThrowIfCancellationRequested();
            return value;
        }

        public override async ValueTask<int> ReadAsync(
            Memory<byte> buffer, CancellationToken ignored = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return read;
        }

        public override async Task<int> ReadAsync(
            byte[] buffer, int offset, int count, CancellationToken ignored)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = await inner.ReadAsync(buffer.AsMemory(offset, count), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                inner.Dispose();
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            await inner.DisposeAsync().ConfigureAwait(false);
            GC.SuppressFinalize(this);
        }
    }

    private static List<OpcRelationship> ParseRelationships(
        Stream stream,
        CancellationToken cancellationToken)
    {
        var relationships = new List<OpcRelationship>();
        using var reader = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
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
