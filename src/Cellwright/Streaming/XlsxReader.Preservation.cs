using Cellwright.IO;
using Cellwright.Styles;

namespace Cellwright;

public sealed partial class XlsxReader
{
    /// <summary>Original part path of a sheet, used to save it back where it came from.</summary>
    internal string SheetPartPath(int sheetIndex) => _sheetParts[sheetIndex];

    /// <summary>Relationship id the workbook used to point at a sheet.</summary>
    internal string SheetRelationshipId(int sheetIndex) => _sheetRelationshipIds[sheetIndex];

    /// <summary>Relationships declared by a sheet part; empty when the sheet has none.</summary>
    internal ValueTask<IReadOnlyList<OpcRelationship>> GetSheetRelationshipsAsync(
        int sheetIndex, CancellationToken cancellationToken) =>
        _package.GetRelationshipsAsync(_sheetParts[sheetIndex], cancellationToken);

    /// <summary>
    /// The open package, for the loaders that walk a part graph of their own rather than
    /// streaming a sheet — drawings and the parts they point at.
    /// </summary>
    internal OpcPackage Package => _package;

    /// <summary>Part name of the workbook, which its own relationships are relative to.</summary>
    internal string WorkbookPartPath => _workbookPart;

    /// <summary>Relationships declared by the workbook part, for the parts that hang off it.</summary>
    internal ValueTask<IReadOnlyList<OpcRelationship>> GetWorkbookRelationshipsAsync(
        CancellationToken cancellationToken) =>
        _package.GetRelationshipsAsync(_workbookPart, cancellationToken);

    /// <summary>Reads a part of the package into memory; used for parts the model owns but does not stream.</summary>
    internal ValueTask<byte[]> ReadPartBytesAsync(string partPath, CancellationToken cancellationToken) =>
        _package.ReadPartBytesAsync(partPath, cancellationToken);

    /// <summary>Whether the package contains the given part.</summary>
    internal bool PartExists(string partPath) => _package.PartExists(partPath);

    /// <summary>
    /// Snapshots everything the document model does not represent: the bytes of unmodelled
    /// parts, the relationship graph and the content types. The workbook keeps this around so
    /// that saving reproduces the parts of the file we never understood in the first place.
    /// </summary>
    internal async ValueTask<PreservedPackage> CreatePreservationAsync(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ContentTypeMap contentTypes = await _package.GetContentTypesAsync(cancellationToken).ConfigureAwait(false);
        StylesPart styles = await GetStylesAsync(cancellationToken).ConfigureAwait(false);
        var preserved = new PreservedPackage
        {
            ContentTypes = contentTypes,
            WorkbookPartPath = _workbookPart,
            WorkbookContentType = contentTypes.GetContentType(_workbookPart) ?? XlsxSchema.ContentTypeWorkbook,
            StylesPartPath = _stylesPart,
            SharedStringsPartPath = _sharedStringsPart,
            StyleMarkup = styles.Markup,
            ForeignSheets = _foreignSheets,
            CorePropertiesPartPath = _corePropertiesPart,
            ExtendedPropertiesPartPath = _extendedPropertiesPart,
            CustomPropertiesPartPath = _customPropertiesPart,
        };

        CopyMarkup(preserved.Markup);

        var regenerated = new HashSet<string>(OpcNames.Instance) { _workbookPart };
        foreach (string sheetPart in _sheetParts)
            regenerated.Add(sheetPart);
        if (_stylesPart is not null)
            regenerated.Add(_stylesPart);
        if (_sharedStringsPart is not null)
            regenerated.Add(_sharedStringsPart);
        if (_calcChainPart is not null)
            regenerated.Add(_calcChainPart);
        foreach (string? propertyPart in new[] { _corePropertiesPart, _extendedPropertiesPart, _customPropertiesPart })
        {
            if (propertyPart is not null)
                regenerated.Add(propertyPart);
        }

        foreach (string partPath in _package.PartPaths.ToArray())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (partPath == "/[Content_Types].xml")
                continue;

            if (OpcPath.IsRelsPath(partPath))
            {
                await CollectRelationshipsAsync(preserved, partPath, cancellationToken).ConfigureAwait(false);
                continue;
            }

            if (regenerated.Contains(partPath))
                continue;

            preserved.Parts[partPath] = await _package.ReadPartBytesAsync(partPath, cancellationToken).ConfigureAwait(false);
        }

        preserved.OrphanParts = FindOrphans(preserved);
        return preserved;
    }

    /// <summary>Parts that were already unreachable in the source package, before any editing.</summary>
    private HashSet<string> FindOrphans(PreservedPackage preserved)
    {
        List<string> roots = ["/", _workbookPart, .. _sheetParts];
        HashSet<string> anchored = preserved.Reachable(roots);

        var orphans = new HashSet<string>(OpcNames.Instance);
        foreach (string partPath in preserved.Parts.Keys)
        {
            if (!anchored.Contains(partPath))
                orphans.Add(partPath);
        }

        return orphans;
    }

    /// <summary>
    /// Records the relationships of one part, dropping the ones the writer produces itself so
    /// that reachability analysis does not treat a deleted sheet's dependencies as live.
    /// </summary>
    private async ValueTask CollectRelationshipsAsync(
        PreservedPackage preserved, string relsPath, CancellationToken cancellationToken)
    {
        string source = OpcPath.GetSourcePart(relsPath);
        IReadOnlyList<OpcRelationship> relationships = await _package.GetRelationshipsAsync(source, cancellationToken)
            .ConfigureAwait(false);
        if (relationships.Count == 0)
            return;

        // A relationship pointing at a part the package never contained makes the result
        // unopenable, and some producers leave those behind. Copying the rot is not fidelity.
        relationships = [.. relationships.Where(r => Resolves(source, r))];

        if (source == _workbookPart)
            relationships = [.. relationships.Where(static r => !IsGeneratedWorkbookRelationship(r.Type))];
        else if (source == "/")
            relationships = [.. relationships.Where(static r => !IsGeneratedRootRelationship(r))];

        if (relationships.Count > 0)
            preserved.Relationships[source] = relationships;
    }

    private bool Resolves(string source, OpcRelationship relationship)
    {
        // Producers routinely omit TargetMode on hyperlinks, whose targets are URLs or
        // in-document anchors rather than parts. Only genuine part references have to exist.
        if (relationship.IsExternal ||
            relationship.Type == XlsxSchema.RelHyperlink ||
            relationship.Target.StartsWith('#') ||
            Uri.TryCreate(relationship.Target, UriKind.Absolute, out _))
        {
            return true;
        }

        return _package.PartExists(OpcPath.Resolve(source, relationship.Target));
    }

    private static bool IsGeneratedWorkbookRelationship(string type) =>
        type is XlsxSchema.RelWorksheet or XlsxSchema.RelStyles or XlsxSchema.RelSharedStrings or XlsxSchema.RelCalcChain;

    /// <summary>
    /// Whether the writer produces this root relationship itself. The document-property parts
    /// are model-owned now, and the target is checked as well as the type: a producer that
    /// spells the type its own way would otherwise leave a second relationship pointing at the
    /// same part, which no consumer opens.
    /// </summary>
    private static bool IsGeneratedRootRelationship(OpcRelationship relationship) =>
        relationship.Type is XlsxSchema.RelCoreProperties or XlsxSchema.RelExtendedProperties or XlsxSchema.RelCustomProperties ||
        (!relationship.IsExternal && XlsxSchema.IsDocumentPropertiesPart(OpcPath.Resolve("/", relationship.Target)));

    private void CopyMarkup(WorkbookMarkup destination)
    {
        destination.RootAttributes = _markup.RootAttributes;
        destination.WorkbookProperties = _markup.WorkbookProperties;
        destination.CalculationProperties = _markup.CalculationProperties;
        destination.BeforeProperties = _markup.BeforeProperties;
        destination.AfterProperties = _markup.AfterProperties;
        destination.BeforeDefinedNames = _markup.BeforeDefinedNames;
        destination.AfterCalculationProperties = _markup.AfterCalculationProperties;
    }
}
