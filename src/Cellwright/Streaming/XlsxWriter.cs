using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright;

/// <summary>
/// Forward-only xlsx writer with constant memory usage: rows are serialized straight to the
/// package as UTF-8, nothing is retained except the shared-string table. Begin a sheet,
/// write rows, dispose. Disposing the writer finalizes the workbook.
/// </summary>
public sealed class XlsxWriter : IAsyncDisposable
{
    private readonly OpcPackage _package;
    private readonly PreservedPackage? _preserved;
    private readonly RelationshipIdAllocator _workbookRelationshipIds = new();
    private readonly List<(string Type, string PartPath)> _addedWorkbookRelationships = [];
    private readonly HashSet<string> _replacedWorkbookRelationshipTypes = new(StringComparer.Ordinal);
    private readonly HashSet<string> _usedPartPaths = new(OpcNames.Instance);
    private readonly HashSet<string> _releasedPreservedPartPaths = new(OpcNames.Instance);
    private readonly List<SheetEntry> _sheetEntries = [];
    private readonly List<DefinedNameEntry> _definedNames = [];
    private readonly string _workbookPartPath;
    private readonly string _workbookContentType;
    private readonly string _stylesPartPath;
    private readonly string _sharedStringsPartPath;
    private readonly string _corePropertiesPartPath;
    private readonly string _extendedPropertiesPartPath;
    private readonly string _customPropertiesPartPath;
    private XlsxSheetWriter? _currentSheet;
    private bool _hasFormulas;
    private bool _disposed;

    private XlsxWriter(OpcPackage package, XlsxWriterOptions options, StylePool? styles, PreservedPackage? preserved)
    {
        _package = package;
        _preserved = preserved;
        Options = options;
        Styles = styles ?? new StylePool();
        DifferentialStyles = new DifferentialStylePool(preserved?.StyleMarkup?.DifferentialFormatCount ?? 0);

        _workbookPartPath = preserved?.WorkbookPartPath ?? XlsxSchema.PartWorkbook;
        _workbookContentType = preserved?.WorkbookContentType ?? XlsxSchema.ContentTypeWorkbook;
        _stylesPartPath = preserved?.StylesPartPath ?? XlsxSchema.PartStyles;
        _sharedStringsPartPath = preserved?.SharedStringsPartPath ?? XlsxSchema.PartSharedStrings;
        _corePropertiesPartPath = preserved?.CorePropertiesPartPath ?? XlsxSchema.PartCoreProperties;
        _extendedPropertiesPartPath = preserved?.ExtendedPropertiesPartPath ?? XlsxSchema.PartExtendedProperties;
        _customPropertiesPartPath = preserved?.CustomPropertiesPartPath ?? XlsxSchema.PartCustomProperties;

        if (preserved is null)
            return;

        _package.MergeContentTypes(preserved.ContentTypes);
        _workbookRelationshipIds.Reserve(preserved.RelationshipsOf(_workbookPartPath));
        foreach (string partPath in preserved.Parts.Keys)
            _usedPartPaths.Add(partPath);
        _usedPartPaths.Add(_workbookPartPath);
        _usedPartPaths.Add(_stylesPartPath);
        _usedPartPaths.Add(_sharedStringsPartPath);
    }

    internal XlsxWriterOptions Options { get; }

    internal SharedStringTableBuilder SharedStrings { get; } = new();

    internal StylePool Styles { get; }

    /// <summary>Differential formats the sheets' conditional formatting rules asked for.</summary>
    internal DifferentialStylePool DifferentialStyles { get; }

    /// <summary>Protection of the sheet list and window, written into the workbook part.</summary>
    internal WorkbookProtection Protection { get; set; } = new();

    /// <summary>What the package will say about itself, written into the parts under <c>docProps</c>.</summary>
    internal WorkbookProperties Properties { get; set; } = new();

    /// <summary>
    /// Declares a defined name. Sheet-scoped names take the zero-based index of a sheet
    /// that must be written by this writer; workbook-scoped names pass <see langword="null"/>.
    /// </summary>
    public void AddDefinedName(string name, string refersTo, int? localSheetIndex = null, bool hidden = false) =>
        AddDefinedName(name, refersTo, localSheetIndex, hidden, attributes: null);

    /// <summary>
    /// Declares a defined name, keeping the attributes of the element it was read from:
    /// its comment, shortcut key and macro flags all live there and nowhere else.
    /// </summary>
    /// <param name="name">Name used in formulas.</param>
    /// <param name="refersTo">Target formula.</param>
    /// <param name="localSheetIndex">Zero-based sheet index for a sheet-scoped name.</param>
    /// <param name="hidden">Whether the name manager hides it.</param>
    /// <param name="attributes">Attributes of a loaded element the model does not represent.</param>
    internal void AddDefinedName(
        string name, string refersTo, int? localSheetIndex, bool hidden, List<KeyValuePair<string, string>>? attributes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrEmpty(refersTo);
        ObjectDisposedException.ThrowIf(_disposed, this);
        _definedNames.Add(new DefinedNameEntry(
            name, refersTo.StartsWith('=') ? refersTo[1..] : refersTo, localSheetIndex, hidden, attributes));
    }

    /// <summary>
    /// Interns a style and returns its handle. Equal styles share one handle, so registering
    /// the same record repeatedly is free. Pass the handle to the style-aware
    /// <see cref="XlsxSheetWriter.WriteRow(ReadOnlySpan{CellValue}, ReadOnlySpan{StyleId})"/> overloads.
    /// </summary>
    public StyleId RegisterStyle(CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        ObjectDisposedException.ThrowIf(_disposed, this);
        return new StyleId(Styles.Intern(style));
    }

    /// <summary>Creates a writer over a stream.</summary>
    public static ValueTask<XlsxWriter> CreateAsync(
        Stream stream, XlsxWriterOptions? options = null, CancellationToken cancellationToken = default)
        => CreateAsync(stream, options, sharedStyles: null, preserved: null, cancellationToken);

    /// <summary>Creates a writer that optionally shares a style pool and carries preserved package content.</summary>
    internal static async ValueTask<XlsxWriter> CreateAsync(
        Stream stream,
        XlsxWriterOptions? options,
        StylePool? sharedStyles,
        PreservedPackage? preserved,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(stream);
        XlsxWriterOptions effective = options ?? XlsxWriterOptions.Default;
        OpcPackage package = await OpcPackage.CreateAsync(stream, effective.LeaveOpen, cancellationToken).ConfigureAwait(false);
        return new XlsxWriter(package, effective, sharedStyles, preserved);
    }

    /// <summary>Creates a writer that produces a file, overwriting any existing one.</summary>
    public static ValueTask<XlsxWriter> CreateAsync(
        string path, XlsxWriterOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var stream = new FileStream(
            path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 64 * 1024, useAsync: true);
        return CreateAsync(stream, (options ?? XlsxWriterOptions.Default) with { LeaveOpen = false }, cancellationToken);
    }

    /// <summary>
    /// Claims part paths and relationship ids that later sheets will reuse from a loaded
    /// package, so that sheets added in this session do not land on top of them.
    /// </summary>
    internal void ReserveLayout(IEnumerable<string> partPaths, IEnumerable<string> relationshipIds)
    {
        foreach (string partPath in partPaths)
            _usedPartPaths.Add(partPath);
        foreach (string id in relationshipIds)
            _workbookRelationshipIds.Reserve(id);
    }

    /// <summary>
    /// Starts the next worksheet. Any previously started sheet is finalized first,
    /// because the package is written strictly forward.
    /// </summary>
    public async ValueTask<XlsxSheetWriter> BeginSheetAsync(
        string name, SheetOptions? options = null, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        SheetNameValidator.Validate(name);
        foreach (SheetEntry existing in _sheetEntries)
        {
            if (string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"A sheet named '{name}' already exists in this workbook.", nameof(name));
        }

        if (_currentSheet is not null)
            await _currentSheet.DisposeAsync().ConfigureAwait(false);

        SheetOptions effective = options ?? SheetOptions.Default;
        string partPath = effective.PartPath ?? NextWorksheetPartPath();
        string relationshipId = effective.RelationshipId ?? _workbookRelationshipIds.Next();
        _usedPartPaths.Add(partPath);
        _workbookRelationshipIds.Reserve(relationshipId);
        _sheetEntries.Add(new SheetEntry(name, effective.Visibility, partPath, relationshipId, effective.SheetId));

        if (effective.Relationships is { Count: > 0 } relationships)
        {
            foreach (OpcRelationship relationship in relationships)
                _package.AddRelationship(partPath, relationship);
        }

        Stream partStream = await _package.CreatePartAsync(partPath, XlsxSchema.ContentTypeWorksheet, cancellationToken)
            .ConfigureAwait(false);

        _currentSheet = new XlsxSheetWriter(this, new Utf8XmlWriter(partStream), effective);
        return _currentSheet;
    }

    internal void OnSheetCompleted(XlsxSheetWriter sheet)
    {
        if (ReferenceEquals(_currentSheet, sheet))
            _currentSheet = null;
    }

    internal void OnFormulaWritten() => _hasFormulas = true;

    /// <summary>
    /// Writes an extra part of the package, such as a comments or table part. Only legal
    /// between sheets: the package is a forward-only zip and holds one open entry at a time.
    /// </summary>
    internal async ValueTask WritePartAsync(
        string partPath, string contentType, Action<Utf8XmlWriter> write, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_currentSheet is not null)
            await _currentSheet.DisposeAsync().ConfigureAwait(false);

        _usedPartPaths.Add(partPath);
        Stream stream = await _package.CreatePartAsync(partPath, contentType, cancellationToken).ConfigureAwait(false);
        await using var xml = new Utf8XmlWriter(stream, bufferSize: 8 * 1024);
        write(xml);
    }

    /// <summary>
    /// Writes a part that is not XML — an image. Same rule as <see cref="WritePartAsync"/>:
    /// only between sheets, because the package holds one open entry at a time.
    /// </summary>
    internal async ValueTask WriteBinaryPartAsync(
        string partPath, string contentType, ReadOnlyMemory<byte> content, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_currentSheet is not null)
            await _currentSheet.DisposeAsync().ConfigureAwait(false);

        _usedPartPaths.Add(partPath);
        Stream stream = await _package.CreatePartAsync(partPath, contentType, cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            await stream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Records the relationships a part of its own declares, e.g. a drawing's images.</summary>
    internal void AddRelationships(string partPath, IReadOnlyList<OpcRelationship> relationships)
    {
        foreach (OpcRelationship relationship in relationships)
            _package.AddRelationship(partPath, relationship);
    }

    /// <summary>Returns a part path of the given shape that no other part in the package occupies.</summary>
    /// <summary>
    /// Registers a part of the workbook itself, replacing whatever the loaded file had under the
    /// same relationship type.
    /// </summary>
    /// <param name="type">The relationship type, which is what identifies the part's role.</param>
    /// <param name="partPath">Where the part was written.</param>
    /// <remarks>
    /// Separate from writing the part because the workbook's relationships are emitted when the
    /// package is finalized, long after the part itself has to be written — a package holds one
    /// open entry at a time.
    /// </remarks>
    internal void AddWorkbookPart(string type, string partPath)
    {
        _replacedWorkbookRelationshipTypes.Add(type);
        _addedWorkbookRelationships.Add((type, partPath));
    }

    /// <summary>Drops the loaded file's part of a role without putting one in its place.</summary>
    /// <param name="type">The relationship type to drop.</param>
    internal void DropWorkbookPart(string type) => _replacedWorkbookRelationshipTypes.Add(type);

    internal string ClaimPartPath(Func<int, string> candidate)
    {
        for (int index = 1; ; index++)
        {
            string partPath = candidate(index);
            if (_usedPartPaths.Add(partPath))
                return partPath;
        }
    }

    /// <summary>
    /// Releases a model-owned source part before regeneration. A canonical source name can then
    /// be reused, and a noncanonical one is omitted instead of surviving as unreachable stale
    /// data in the saved package.
    /// </summary>
    internal void ReleasePreservedPart(string partPath)
    {
        if (_preserved?.Parts.ContainsKey(partPath) != true)
            return;

        _releasedPreservedPartPaths.Add(partPath);
        _usedPartPaths.Remove(partPath);
    }

    /// <summary>Finalizes the workbook: shared strings, styles, workbook part, relationships, content types.</summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        if (_currentSheet is not null)
        {
            await _currentSheet.DisposeAsync().ConfigureAwait(false);
            _currentSheet = null;
        }

        // A workbook must contain at least one sheet, but a preserved chart or dialog sheet is
        // already one. Only a genuinely empty new workbook needs a generated worksheet.
        if (_sheetEntries.Count == 0 && (_preserved?.ForeignSheets.Count ?? 0) == 0)
        {
            XlsxSheetWriter empty = await BeginSheetAsync("Sheet1").ConfigureAwait(false);
            await empty.DisposeAsync().ConfigureAwait(false);
        }

        _disposed = true;
        bool hasSharedStrings = SharedStrings.UniqueCount > 0;
        if (hasSharedStrings)
            await WriteSharedStringsAsync().ConfigureAwait(false);
        await WriteStylesAsync().ConfigureAwait(false);
        await WriteWorkbookAsync().ConfigureAwait(false);
        await WritePropertiesAsync().ConfigureAwait(false);
        await WritePreservedPartsAsync().ConfigureAwait(false);
        AddRelationships(hasSharedStrings);
        await _package.DisposeAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Writes the parts under <c>docProps</c>. Every package gets the first two, as Excel's own
    /// do; the third only exists when somebody put something in it.
    /// </summary>
    private async ValueTask WritePropertiesAsync()
    {
        List<string> sheetNames = [.. _sheetEntries.Select(static sheet => sheet.Name)];
        await WriteFinalPartAsync(
            _corePropertiesPartPath,
            XlsxSchema.ContentTypeCoreProperties,
            xml => DocumentPropertiesPart.WriteCore(xml, Properties)).ConfigureAwait(false);
        await WriteFinalPartAsync(
            _extendedPropertiesPartPath,
            XlsxSchema.ContentTypeExtendedProperties,
            xml => DocumentPropertiesPart.WriteExtended(xml, Properties, sheetNames)).ConfigureAwait(false);

        if (HasCustomProperties)
        {
            await WriteFinalPartAsync(
                _customPropertiesPartPath,
                XlsxSchema.ContentTypeCustomProperties,
                xml => DocumentPropertiesPart.WriteCustom(xml, Properties)).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Whether to write the third part. A package that arrived with an empty one keeps it:
    /// removing a part nobody asked to remove is a change, however meaningless its content.
    /// </summary>
    private bool HasCustomProperties => Properties.Custom.Count > 0 || _preserved?.CustomPropertiesPartPath is not null;

    /// <summary>Writes a part while finalizing, when <see cref="WritePartAsync"/> is already closed for business.</summary>
    private async ValueTask WriteFinalPartAsync(string partPath, string contentType, Action<Utf8XmlWriter> write)
    {
        _usedPartPaths.Add(partPath);
        Stream stream = await _package.CreatePartAsync(partPath, contentType, CancellationToken.None).ConfigureAwait(false);
        await using var xml = new Utf8XmlWriter(stream, bufferSize: 8 * 1024);
        write(xml);
    }

    private string NextWorksheetPartPath() => ClaimPartPath(XlsxSchema.WorksheetPart);

    private async ValueTask WriteSharedStringsAsync()
    {
        Stream stream = await _package.CreatePartAsync(_sharedStringsPartPath, XlsxSchema.ContentTypeSharedStrings, CancellationToken.None)
            .ConfigureAwait(false);
        await using var xml = new Utf8XmlWriter(stream);

        xml.WriteDeclaration();
        xml.WriteRaw("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\""u8);
        xml.WriteFormattable(SharedStrings.ReferenceCount);
        xml.WriteRaw("\" uniqueCount=\""u8);
        xml.WriteInt32(SharedStrings.UniqueCount);
        xml.WriteRaw("\">"u8);

        foreach (SharedStringItem item in SharedStrings.Items)
        {
            XlsxSheetWriter.WriteSharedStringItem(xml, item.Text, item.RichMarkup);
            await xml.FlushIfNeededAsync(CancellationToken.None).ConfigureAwait(false);
        }

        xml.WriteRaw("</sst>"u8);
    }

    private async ValueTask WriteStylesAsync()
    {
        Stream stream = await _package.CreatePartAsync(_stylesPartPath, XlsxSchema.ContentTypeStyles, CancellationToken.None)
            .ConfigureAwait(false);
        await using var xml = new Utf8XmlWriter(stream, bufferSize: 16 * 1024);
        StylesPartWriter.Write(xml, Styles.Styles, _preserved?.StyleMarkup, DifferentialStyles.Styles);
    }

    /// <summary>Copies the parts of a loaded package that no sheet or workbook element regenerates.</summary>
    private async ValueTask WritePreservedPartsAsync()
    {
        if (_preserved is null)
            return;

        var roots = new List<string>(_sheetEntries.Count + 2) { "/", _workbookPartPath };
        foreach (SheetEntry sheet in _sheetEntries)
            roots.Add(sheet.PartPath);

        HashSet<string> reachable = _preserved.Reachable(roots);
        foreach ((string partPath, byte[] content) in _preserved.Parts)
        {
            if (_releasedPreservedPartPaths.Contains(partPath))
                continue;
            if (!reachable.Contains(partPath) && !_preserved.OrphanParts.Contains(partPath))
                continue;

            // A package whose document properties nothing pointed at keeps them among the
            // orphans; the generated ones are the ones that belong at that name now.
            if (_package.PartWritten(partPath))
                continue;

            await _package.WriteRawPartAsync(partPath, content, CancellationToken.None).ConfigureAwait(false);
        }

        // The copied parts point at each other (a drawing at its images, a pivot table at its
        // cache); their relationship parts have to travel with them.
        foreach ((string source, IReadOnlyList<OpcRelationship> relationships) in _preserved.Relationships)
        {
            if (source == "/" || source == _workbookPartPath)
                continue;
            if (_releasedPreservedPartPaths.Contains(source))
                continue;
            if (!reachable.Contains(source) && !_preserved.OrphanParts.Contains(source))
                continue;

            foreach (OpcRelationship relationship in relationships)
                _package.AddRelationship(source, relationship);
        }
    }

    private async ValueTask WriteWorkbookAsync()
    {
        Stream stream = await _package.CreatePartAsync(_workbookPartPath, _workbookContentType, CancellationToken.None)
            .ConfigureAwait(false);
        await using var xml = new Utf8XmlWriter(stream, bufferSize: 8 * 1024);
        WorkbookMarkup? markup = _preserved?.Markup;

        xml.WriteDeclaration();
        xml.WriteRaw("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""u8);
        xml.WriteRaw(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""u8);
        RawXml.WriteAttributes(xml, markup?.RootAttributes);
        xml.WriteRaw(">"u8);

        RawXml.Write(xml, markup?.BeforeProperties);
        WriteWorkbookProperties(xml, markup);
        ProtectionPart.WriteWorkbook(xml, Protection);
        RawXml.Write(xml, markup?.AfterProperties);

        WriteSheetList(xml);
        RawXml.Write(xml, markup?.BeforeDefinedNames);
        WriteDefinedNames(xml);
        WriteCalculationProperties(xml, markup);
        RawXml.Write(xml, markup?.AfterCalculationProperties);
        xml.WriteRaw("</workbook>"u8);
    }

    /// <summary>
    /// Writes the sheet list, slotting chart and dialog sheets back where they were. Their
    /// tabs would otherwise vanish from a workbook that only had a cell edited.
    /// </summary>
    private void WriteSheetList(Utf8XmlWriter xml)
    {
        IReadOnlyList<ForeignSheet> foreign = _preserved?.ForeignSheets ?? [];

        // Sheets that came from a file keep the identifier they had; only new ones get a number,
        // taken from above every identifier already in use.
        int nextId = 1;
        foreach (SheetEntry worksheet in _sheetEntries)
            nextId = Math.Max(nextId, worksheet.SheetId + 1);
        foreach (ForeignSheet sheet in foreign)
            nextId = Math.Max(nextId, sheet.SheetId + 1);

        int written = 0;
        int nextForeign = 0;
        xml.WriteRaw("<sheets>"u8);
        foreach (SheetEntry worksheet in _sheetEntries)
        {
            while (nextForeign < foreign.Count && foreign[nextForeign].Position <= written)
            {
                ForeignSheet sheet = foreign[nextForeign++];
                WriteSheetEntry(xml, sheet.Name, sheet.State, IdOf(sheet.SheetId, ref nextId), sheet.RelationshipId);
                written++;
            }

            WriteSheetEntry(
                xml, worksheet.Name, StateOf(worksheet.Visibility),
                IdOf(worksheet.SheetId, ref nextId), worksheet.RelationshipId);
            written++;
        }

        while (nextForeign < foreign.Count)
        {
            ForeignSheet sheet = foreign[nextForeign++];
            WriteSheetEntry(xml, sheet.Name, sheet.State, IdOf(sheet.SheetId, ref nextId), sheet.RelationshipId);
        }

        xml.WriteRaw("</sheets>"u8);

        static int IdOf(int preserved, ref int nextId) => preserved > 0 ? preserved : nextId++;
    }

    /// <summary>The <c>ST_SheetState</c> value, or null for a visible sheet, which writes none.</summary>
    private static string? StateOf(SheetVisibility visibility) => visibility switch
    {
        SheetVisibility.Hidden => "hidden",
        SheetVisibility.VeryHidden => "veryHidden",
        _ => null,
    };

    private static void WriteSheetEntry(Utf8XmlWriter xml, string name, string? state, int sheetId, string relationshipId)
    {
        xml.WriteRaw("<sheet name=\""u8);
        xml.WriteAttributeText(name);
        xml.WriteRaw("\" sheetId=\""u8);
        xml.WriteInt32(sheetId);
        if (state is not null)
        {
            xml.WriteRaw("\" state=\""u8);
            xml.WriteAttributeText(state);
        }

        xml.WriteRaw("\" r:id=\""u8);
        xml.WriteAttributeText(relationshipId);
        xml.WriteRaw("\"/>"u8);
    }

    private void WriteWorkbookProperties(Utf8XmlWriter xml, WorkbookMarkup? markup)
    {
        bool date1904 = Options.DateSystem == ExcelDateSystem.Date1904;
        if (!date1904 && markup?.WorkbookProperties is null)
            return;

        xml.WriteRaw("<workbookPr"u8);
        RawXml.WriteAttributes(xml, markup?.WorkbookProperties);
        if (date1904)
            xml.WriteRaw(" date1904=\"1\""u8);
        xml.WriteRaw("/>"u8);
    }

    private void WriteCalculationProperties(Utf8XmlWriter xml, WorkbookMarkup? markup)
    {
        bool fullCalcOnLoad = _hasFormulas && Options.FullCalculationOnLoad;
        if (!fullCalcOnLoad && markup?.CalculationProperties is null)
            return;

        xml.WriteRaw("<calcPr"u8);
        RawXml.WriteAttributes(xml, markup?.CalculationProperties);
        if (fullCalcOnLoad)
            xml.WriteRaw(" fullCalcOnLoad=\"1\""u8);
        xml.WriteRaw("/>"u8);
    }

    private void WriteDefinedNames(Utf8XmlWriter xml)
    {
        if (_definedNames.Count == 0)
            return;

        xml.WriteRaw("<definedNames>"u8);
        foreach ((string name, string refersTo, int? localSheetIndex, bool hidden, var attributes) in _definedNames)
        {
            xml.WriteRaw("<definedName name=\""u8);
            xml.WriteAttributeText(name);
            xml.WriteRaw("\""u8);
            if (localSheetIndex is { } index)
            {
                xml.WriteRaw(" localSheetId=\""u8);
                xml.WriteInt32(FullSheetPosition(index));
                xml.WriteRaw("\""u8);
            }

            if (hidden)
                xml.WriteRaw(" hidden=\"1\""u8);

            RawXml.WriteAttributes(xml, attributes);
            xml.WriteRaw(">"u8);
            xml.WriteText(FutureFunctions.QualifyAll(refersTo)!);
            xml.WriteRaw("</definedName>"u8);
        }

        xml.WriteRaw("</definedNames>"u8);
    }

    /// <summary>
    /// Converts the worksheet-only index used by the object model back to the position in the
    /// complete workbook sheet list, where chart, macro and dialog sheets also take a slot.
    /// </summary>
    private int FullSheetPosition(int worksheetIndex)
    {
        int position = worksheetIndex;
        foreach (ForeignSheet foreign in _preserved?.ForeignSheets ?? [])
        {
            if (foreign.Position > position)
                break;
            position++;
        }

        return position;
    }

    private void AddRelationships(bool hasSharedStrings)
    {
        IReadOnlyList<OpcRelationship> preservedRoot = _preserved?.RelationshipsOf("/") ?? [];
        var rootIds = new RelationshipIdAllocator();
        rootIds.Reserve(preservedRoot);

        bool hasOfficeDocument = false;
        foreach (OpcRelationship relationship in preservedRoot)
        {
            _package.AddRelationship("/", relationship);
            hasOfficeDocument |= relationship.Type == XlsxSchema.RelOfficeDocument;
        }

        if (!hasOfficeDocument)
        {
            _package.AddRelationship("/", new OpcRelationship(
                rootIds.Next(), XlsxSchema.RelOfficeDocument, OpcPath.MakeRelative("/", _workbookPartPath)));
        }

        AddRootRelationship(rootIds, XlsxSchema.RelCoreProperties, _corePropertiesPartPath);
        AddRootRelationship(rootIds, XlsxSchema.RelExtendedProperties, _extendedPropertiesPartPath);
        if (HasCustomProperties)
            AddRootRelationship(rootIds, XlsxSchema.RelCustomProperties, _customPropertiesPartPath);

        foreach (OpcRelationship relationship in _preserved?.RelationshipsOf(_workbookPartPath) ?? [])
        {
            // A part the model regenerated replaces the one the file had: keeping both would
            // point the workbook at two lists of the same thing.
            if (_replacedWorkbookRelationshipTypes.Contains(relationship.Type))
                continue;

            _package.AddRelationship(_workbookPartPath, relationship);
        }

        foreach ((string type, string partPath) in _addedWorkbookRelationships)
        {
            _package.AddRelationship(_workbookPartPath, new OpcRelationship(
                _workbookRelationshipIds.Next(), type, OpcPath.MakeRelative(_workbookPartPath, partPath)));
        }

        foreach (SheetEntry sheet in _sheetEntries)
        {
            _package.AddRelationship(_workbookPartPath, new OpcRelationship(
                sheet.RelationshipId, XlsxSchema.RelWorksheet, OpcPath.MakeRelative(_workbookPartPath, sheet.PartPath)));
        }

        _package.AddRelationship(_workbookPartPath, new OpcRelationship(
            _workbookRelationshipIds.Next(), XlsxSchema.RelStyles, OpcPath.MakeRelative(_workbookPartPath, _stylesPartPath)));
        if (hasSharedStrings)
        {
            _package.AddRelationship(_workbookPartPath, new OpcRelationship(
                _workbookRelationshipIds.Next(), XlsxSchema.RelSharedStrings,
                OpcPath.MakeRelative(_workbookPartPath, _sharedStringsPartPath)));
        }
    }

    private void AddRootRelationship(RelationshipIdAllocator ids, string type, string partPath) =>
        _package.AddRelationship("/", new OpcRelationship(ids.Next(), type, OpcPath.MakeRelative("/", partPath)));

    private readonly record struct SheetEntry(
        string Name, SheetVisibility Visibility, string PartPath, string RelationshipId, int SheetId);

    private readonly record struct DefinedNameEntry(
        string Name,
        string RefersTo,
        int? LocalSheetIndex,
        bool Hidden,
        List<KeyValuePair<string, string>>? Attributes);
}
