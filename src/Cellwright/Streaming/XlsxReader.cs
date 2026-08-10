using System.Globalization;
using System.Text;
using System.Xml;
using Cellwright.IO;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright;

/// <summary>
/// Forward-only pull reader for xlsx workbooks: enumerate rows of a sheet as
/// <see cref="RowData"/> with constant memory (only the shared-string table is retained).
/// Not thread-safe; enumerate one sheet at a time.
/// </summary>
public sealed partial class XlsxReader : IAsyncDisposable
{
    private readonly OpcPackage _package;
    private readonly XlsxReaderOptions _options;
    private readonly List<SheetInfo> _sheets = [];
    private readonly List<string> _sheetParts = [];
    private readonly List<string> _sheetRelationshipIds = [];
    private readonly List<DefinedNameInfo> _definedNames = [];
    private readonly Dictionary<string, string> _relationshipTypes = new(StringComparer.Ordinal);
    private readonly List<ForeignSheet> _foreignSheets = [];
    private readonly Dictionary<int, int> _sheetPositions = [];
    private readonly WorkbookMarkup _markup = new();
    private int _sheetPosition;
    private string _workbookPart = XlsxSchema.PartWorkbook;
    private string? _sharedStringsPart;
    private string? _stylesPart;
    private string? _calcChainPart;
    private string? _corePropertiesPart;
    private string? _extendedPropertiesPart;
    private string? _customPropertiesPart;
    private string[]? _sharedStrings;
    private string?[]? _richStrings;
    private StylesPart? _styles;
    private TemporalKind[]? _styleTemporalKinds;
    private bool _disposed;

    private XlsxReader(OpcPackage package, XlsxReaderOptions options)
    {
        _package = package;
        _options = options;
    }

    /// <summary>The worksheets of the workbook, in workbook order.</summary>
    public IReadOnlyList<SheetInfo> Sheets => _sheets;

    /// <summary>Defined names declared by the workbook.</summary>
    public IReadOnlyList<DefinedNameInfo> DefinedNames => _definedNames;

    /// <summary>Date system declared by the workbook.</summary>
    public ExcelDateSystem DateSystem { get; private set; }

    /// <summary>Structure protection declared by the workbook.</summary>
    public WorkbookProtection Protection { get; } = new();

    /// <summary>What the package says about itself: title, author, company, custom fields.</summary>
    public WorkbookProperties Properties { get; } = new();

    /// <summary>
    /// Whether the package belongs to the Strict conformance class of ISO/IEC 29500, which
    /// spells every namespace under purl.oclc.org. Its content reads normally; what the writer
    /// produces is always transitional, so the two cannot be mixed in one saved package.
    /// </summary>
    public bool IsStrict { get; private set; }

    /// <summary>Opens a workbook from a stream.</summary>
    public static async ValueTask<XlsxReader> OpenAsync(
        Stream stream, XlsxReaderOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        XlsxReaderOptions effective = options ?? XlsxReaderOptions.Default;
        DocumentLoadBudget budget = effective.EffectiveBudget;
        budget.Validate();
        await FileSignature.RejectNonPackageAsync(stream, cancellationToken).ConfigureAwait(false);
        OpcPackage package = await OpcPackage
            .OpenReadAsync(stream, effective.LeaveOpen, cancellationToken, budget)
            .ConfigureAwait(false);
        return await OpenAsync(package, effective, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Opens a package already admitted under a shared outer load budget.</summary>
    internal static async ValueTask<XlsxReader> OpenAsync(
        OpcPackage package,
        XlsxReaderOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(package);
        ArgumentNullException.ThrowIfNull(options);
        var reader = new XlsxReader(package, options);
        try
        {
            await reader.InitializeAsync(cancellationToken).ConfigureAwait(false);
            return reader;
        }
        catch
        {
            await reader.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    /// <summary>Opens a workbook file.</summary>
    public static ValueTask<XlsxReader> OpenAsync(
        string path, XlsxReaderOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 64 * 1024, useAsync: true);
        return OpenAsync(stream, (options ?? XlsxReaderOptions.Default) with { LeaveOpen = false }, cancellationToken);
    }

    /// <summary>Releases the package and the source stream (unless opened with LeaveOpen).</summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;
        _disposed = true;
        await _package.DisposeAsync().ConfigureAwait(false);
    }

    private async ValueTask InitializeAsync(CancellationToken cancellationToken)
    {
        IReadOnlyList<OpcRelationship> rootRelationships = await _package.GetRelationshipsAsync("/", cancellationToken)
            .ConfigureAwait(false);
        _workbookPart = ResolveWorkbookPart(rootRelationships);
        await ReadPropertiesAsync(rootRelationships, cancellationToken).ConfigureAwait(false);
        Dictionary<string, string> relationshipTargets = await LoadWorkbookRelationshipsAsync(_workbookPart, cancellationToken)
            .ConfigureAwait(false);
        await ParseWorkbookAsync(_workbookPart, relationshipTargets, cancellationToken).ConfigureAwait(false);
    }

    private string ResolveWorkbookPart(IReadOnlyList<OpcRelationship> rootRelationships)
    {
        foreach (OpcRelationship relationship in rootRelationships)
        {
            if (relationship.Type != XlsxSchema.RelOfficeDocument || relationship.IsExternal)
                continue;

            string target = OpcPath.Resolve("/", relationship.Target);
            return ForeignPackage.DescribeWorkbookPart(target) is { } foreign
                ? throw new XlsxFormatException(foreign)
                : target;
        }

        // Recovery: fall back to the canonical location.
        if (_package.PartExists(XlsxSchema.PartWorkbook))
        {
            Warn("MissingOfficeDocumentRel", "The package has no officeDocument relationship; using /xl/workbook.xml.");
            return XlsxSchema.PartWorkbook;
        }

        throw new XlsxFormatException(ForeignPackage.Describe(_package) ?? "The package contains no workbook part.");
    }

    /// <summary>
    /// Reads the parts under <c>docProps</c>. They hang off the package root rather than the
    /// workbook, and a package is free to have none of them.
    /// </summary>
    private async ValueTask ReadPropertiesAsync(
        IReadOnlyList<OpcRelationship> rootRelationships, CancellationToken cancellationToken)
    {
        foreach (OpcRelationship relationship in rootRelationships)
        {
            if (relationship.IsExternal)
                continue;

            Action<Stream, WorkbookProperties>? read = relationship.Type switch
            {
                XlsxSchema.RelCoreProperties => Streaming.DocumentPropertiesPart.ReadCore,
                XlsxSchema.RelExtendedProperties => Streaming.DocumentPropertiesPart.ReadExtended,
                XlsxSchema.RelCustomProperties => Streaming.DocumentPropertiesPart.ReadCustom,
                _ => null,
            };

            string target = OpcPath.Resolve("/", relationship.Target);
            if (read is null || !_package.PartExists(target))
                continue;

            // Packages written through System.IO.Packaging keep their core properties under
            // package/services with a generated name, so saving has to put them back there.
            switch (relationship.Type)
            {
                case XlsxSchema.RelCoreProperties: _corePropertiesPart = target; break;
                case XlsxSchema.RelExtendedProperties: _extendedPropertiesPart = target; break;
                default: _customPropertiesPart = target; break;
            }

            try
            {
                using var buffered = new MemoryStream(
                    await _package.ReadPartBytesAsync(target, cancellationToken).ConfigureAwait(false), writable: false);
                read(buffered, Properties);
            }
            catch (XmlException exception)
            {
                Warn("BadProperties", $"The document properties in '{target}' could not be read: {exception.Message}");
            }
        }
    }

    private async ValueTask<Dictionary<string, string>> LoadWorkbookRelationshipsAsync(
        string workbookPart, CancellationToken cancellationToken)
    {
        IReadOnlyList<OpcRelationship> relationships = await _package.GetRelationshipsAsync(workbookPart, cancellationToken)
            .ConfigureAwait(false);
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.IsExternal)
                continue;

            string target = OpcPath.Resolve(workbookPart, relationship.Target);
            targets[relationship.Id] = target;
            _relationshipTypes[relationship.Id] = relationship.Type;
            if (relationship.Type == XlsxSchema.RelSharedStrings)
                _sharedStringsPart = target;
            else if (relationship.Type == XlsxSchema.RelStyles)
                _stylesPart = target;
            else if (relationship.Type == XlsxSchema.RelCalcChain)
                _calcChainPart = target;
        }

        return targets;
    }

    private async ValueTask ParseWorkbookAsync(
        string workbookPart, Dictionary<string, string> relationshipTargets, CancellationToken cancellationToken)
    {
        Stream stream = await _package.OpenPartAsync(workbookPart, cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            using var xml = XmlReader.Create(stream, XmlDefaults.AsyncReaderSettings);
            StringBuilder? beforeProperties = null;
            StringBuilder? afterProperties = null;
            StringBuilder? beforeDefinedNames = null;
            StringBuilder? trailing = null;
            bool advance = true;
            while (true)
            {
                if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                    break;
                advance = true;

                if (xml.NodeType != XmlNodeType.Element)
                    continue;

                switch (xml.LocalName)
                {
                    case "workbook":
                        _markup.RootAttributes = RawXml.ReadAttributes(xml, "xmlns", "xmlns:r");
                        IsStrict = xml.NamespaceURI.StartsWith("http://purl.oclc.org/ooxml/", StringComparison.Ordinal);
                        WarnAboutMustUnderstand(xml);
                        break;

                    case "workbookPr":
                        if (xml.GetAttribute("date1904") is "1" or "true")
                            DateSystem = ExcelDateSystem.Date1904;
                        _markup.WorkbookProperties = RawXml.ReadAttributes(xml, "date1904");
                        break;

                    case "calcPr":
                        _markup.CalculationProperties = RawXml.ReadAttributes(xml, "fullCalcOnLoad");
                        break;

                    case "sheets" or "definedNames":
                        break;

                    case "sheet":
                        AddSheet(xml, relationshipTargets);
                        break;

                    case "definedName":
                        // Reading the element's text also steps past its end tag, so the loop must
                        // not advance again or the next name would be skipped.
                        advance = xml.IsEmptyElement;
                        await AddDefinedNameAsync(xml).ConfigureAwait(false);
                        break;

                    // Everything else is carried through verbatim in the slot the schema gives it.
                    case "fileVersion" or "fileSharing":
                        beforeProperties = await RawXml.AppendOuterXmlAsync(beforeProperties, xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    case "workbookProtection":
                        Streaming.ProtectionPart.ReadWorkbook(xml, Protection);
                        await xml.SkipAsync().ConfigureAwait(false);
                        advance = false;
                        break;

                    case "bookViews":
                        afterProperties = await RawXml.AppendOuterXmlAsync(afterProperties, xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    case "functionGroups" or "externalReferences":
                        beforeDefinedNames = await RawXml.AppendOuterXmlAsync(beforeDefinedNames, xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    default:
                        trailing = await RawXml.AppendOuterXmlAsync(trailing, xml).ConfigureAwait(false);
                        advance = false;
                        break;
                }
            }

            _markup.BeforeProperties = beforeProperties?.ToString();
            _markup.AfterProperties = afterProperties?.ToString();
            _markup.BeforeDefinedNames = beforeDefinedNames?.ToString();
            _markup.AfterCalculationProperties = trailing?.ToString();
        }

        // A workbook of nothing but chart sheets is valid, and the model simply has no grid to
        // show for it; only a workbook with no tabs at all is broken.
        if (_sheets.Count == 0 && _foreignSheets.Count == 0)
            throw new XlsxFormatException("The workbook declares no sheets.");
    }

    private async ValueTask AddDefinedNameAsync(XmlReader xml)
    {
        string? name = xml.GetAttribute("name");
        int? localSheetId = int.TryParse(
            xml.GetAttribute("localSheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int local)
            ? local
            : null;
        bool hidden = xml.GetAttribute("hidden") is "1" or "true";

        // A defined name can also carry a comment, a shortcut key, and the flags that make it a
        // macro rather than a reference — none of them modelled, all of them worth keeping.
        List<KeyValuePair<string, string>>? attributes = RawXml.ReadAttributes(xml, "name", "localSheetId", "hidden");
        if (xml.IsEmptyElement || name is null)
            return;

        string refersTo = await xml.ReadElementContentAsStringAsync().ConfigureAwait(false);
        if (refersTo.Length == 0)
            return;

        // localSheetId indexes the workbook's whole sheet list, chart sheets included, while
        // the model only holds worksheets. Names scoped to a sheet we do not model are dropped.
        int? scopeIndex = null;
        if (localSheetId is { } position)
        {
            if (!_sheetPositions.TryGetValue(position, out int index))
            {
                Warn("BadNameScope", $"Defined name '{name}' is scoped to a sheet outside the worksheet model and was dropped.");
                return;
            }

            scopeIndex = index;
        }

        _definedNames.Add(
            new DefinedNameInfo(name, FutureFunctions.UnqualifyAll(refersTo)!, scopeIndex, hidden)
            {
                Attributes = attributes,
            });
    }

    private void AddSheet(XmlReader xml, Dictionary<string, string> relationshipTargets)
    {
        string name = xml.GetAttribute("name") ?? $"Sheet{_sheets.Count + 1}";
        string? state = xml.GetAttribute("state");
        string? relationshipId = RelationshipId(xml);
        int sheetId = int.TryParse(
            xml.GetAttribute("sheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int parsedId)
            ? parsedId
            : 0;
        _sheetPosition++;

        if (relationshipId is null || !relationshipTargets.TryGetValue(relationshipId, out string? partPath))
        {
            Warn("SheetWithoutPart", $"Sheet '{name}' has no resolvable worksheet part and was skipped.");
            return;
        }

        // Chart sheets and dialog sheets live in the same list but hold no grid. They are
        // carried through untouched instead of being parsed as worksheets and mangled.
        if (_relationshipTypes.GetValueOrDefault(relationshipId) is { } type && type != XlsxSchema.RelWorksheet)
        {
            _foreignSheets.Add(new ForeignSheet(_sheetPosition - 1, name, state, relationshipId, sheetId));
            return;
        }

        _sheetPositions[_sheetPosition - 1] = _sheets.Count;
        _sheets.Add(new SheetInfo(_sheets.Count, name, ParseVisibility(state), sheetId));
        _sheetParts.Add(partPath);
        _sheetRelationshipIds.Add(relationshipId);
    }

    /// <summary>
    /// Reports the namespaces a file insists its reader understands (29500-3 §9.4). Ignoring
    /// them silently is the one thing the markup compatibility rules forbid: the producer said
    /// the document is only meaningful to a consumer that knows these extensions.
    /// </summary>
    private void WarnAboutMustUnderstand(XmlReader xml)
    {
        if (xml.GetAttribute("MustUnderstand", XlsxSchema.NsMarkupCompatibility) is not { Length: > 0 } prefixes)
            return;

        Warn(
            "MustUnderstand",
            $"The workbook requires a consumer that understands '{prefixes}'. " +
            "Those extensions are carried through unread, so what they describe is not applied.");
    }

    /// <summary>
    /// Reads the <c>r:id</c> attribute. A Strict file puts it in its own namespace, so both are
    /// tried rather than the file being turned away for a missing part.
    /// </summary>
    internal static string? RelationshipId(XmlReader xml) =>
        xml.GetAttribute("id", XlsxSchema.NsDocumentRelationships) ??
        xml.GetAttribute("id", XlsxSchema.NsStrictDocumentRelationships);

    /// <summary>Reads <c>ST_SheetState</c>; anything unrecognized leaves the sheet visible.</summary>
    private static SheetVisibility ParseVisibility(string? state) => state switch
    {
        "hidden" => SheetVisibility.Hidden,
        "veryHidden" => SheetVisibility.VeryHidden,
        _ => SheetVisibility.Visible,
    };

    private async ValueTask EnsureSharedStringsAsync(CancellationToken cancellationToken)
    {
        if (_sharedStrings is not null)
            return;

        if (_sharedStringsPart is null || !_package.PartExists(_sharedStringsPart))
        {
            _sharedStrings = [];
            _richStrings = [];
            return;
        }

        var items = new List<string>();
        var richItems = new List<string?>();
        Stream stream = await _package
            .OpenPartAsync(_sharedStringsPart, cancellationToken, retainedInMemory: true)
            .ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            using var xml = XmlReader.Create(stream, XmlDefaults.AsyncReaderSettings);
            while (await xml.ReadAsync().ConfigureAwait(false))
            {
                if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "si")
                    continue;

                (string text, string? rich) = await ReadSharedItemAsync(xml).ConfigureAwait(false);
                items.Add(text);
                richItems.Add(rich);
            }
        }

        _sharedStrings = [.. items];
        _richStrings = [.. richItems];
    }

    /// <summary>
    /// Reads one shared-string item, keeping the markup of a rich one. Runs are the only
    /// formatting a cell value carries and the model has no room for them, so the original
    /// children are copied out whole and written back for any string that was not edited.
    /// </summary>
    private static async ValueTask<(string Text, string? RichMarkup)> ReadSharedItemAsync(XmlReader xml)
    {
        if (xml.IsEmptyElement)
            return (string.Empty, null);

        var text = new System.Text.StringBuilder();
        System.Text.StringBuilder? markup = null;
        bool hasRuns = false;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "si")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            string name = xml.LocalName;
            if (name == "t")
            {
                text.Append(await xml.ReadElementContentAsStringAsync().ConfigureAwait(false));
                advance = false;
                continue;
            }

            if (name is not ("r" or "rPh" or "phoneticPr"))
                continue;

            string fragment = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
            advance = false;
            (markup ??= new System.Text.StringBuilder()).Append(fragment);

            // Only the runs carry the value; a phonetic hint is a reading aid beside it.
            if (name != "r")
                continue;

            hasRuns = true;
            AppendRunText(text, fragment);
        }

        // Without runs there is nothing the plain path cannot write, and replaying the captured
        // markup on its own would leave a phonetic hint where the text used to be.
        return (OoxmlText.Decode(text.ToString()), hasRuns ? markup?.ToString() : null);
    }

    /// <summary>Pulls the text out of a captured <c>r</c> element.</summary>
    private static void AppendRunText(System.Text.StringBuilder text, string fragment)
    {
        using var reader = XmlReader.Create(new StringReader(fragment), XmlDefaults.ReaderSettings);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
                text.Append(reader.ReadElementContentAsString());
        }
    }

    /// <summary>The parsed styles part; available after any row enumeration started (used by the DOM loader).</summary>
    internal async ValueTask<StylesPart> GetStylesAsync(CancellationToken cancellationToken)
    {
        await EnsureStylesAsync(cancellationToken).ConfigureAwait(false);
        return _styles!;
    }

    private async ValueTask EnsureStylesAsync(CancellationToken cancellationToken)
    {
        if (_styles is not null)
            return;

        if (_stylesPart is null || !_package.PartExists(_stylesPart))
        {
            _styles = StylesPart.Empty;
            _styleTemporalKinds = _styles.TemporalKinds;
            return;
        }

        // Read through the package rather than copied here: the styles part is held whole to be
        // parsed twice, and everything held whole answers to the same budget.
        byte[] content = await _package.ReadPartBytesAsync(_stylesPart, cancellationToken).ConfigureAwait(false);
        using var buffered = new MemoryStream(content, writable: false);
        _styles = StylesPart.Parse(buffered);
        _styleTemporalKinds = _styles.TemporalKinds;
    }

    /// <summary>Reads the concatenated text of a string item (plain <c>t</c> or rich-text runs), decoding OOXML escapes.</summary>
    private static async ValueTask<string> ReadStringItemAsync(XmlReader xml, string containerName)
    {
        if (xml.IsEmptyElement)
            return string.Empty;

        string? single = null;
        System.Text.StringBuilder? builder = null;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == containerName)
                break;

            if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "t")
            {
                string chunk = await xml.ReadElementContentAsStringAsync().ConfigureAwait(false);
                advance = false;
                if (single is null && builder is null)
                {
                    single = chunk;
                }
                else
                {
                    builder ??= new System.Text.StringBuilder(single);
                    single = null;
                    builder.Append(chunk);
                }
            }
            else if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "rPh")
            {
                // Phonetic hints are presentation-only.
                await xml.SkipAsync().ConfigureAwait(false);
                advance = false;
            }
        }

        return OoxmlText.Decode(builder?.ToString() ?? single ?? string.Empty);
    }

    private void Warn(string code, string message, CellRef? cell = null) =>
        _options.OnWarning?.Invoke(new WorkbookWarning(code, message, cell));
}
