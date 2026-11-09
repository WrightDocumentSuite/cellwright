using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Cellwright.IO;

namespace Cellwright.Sanitization;

/// <summary>
/// Removes categories of content from a finished package. Working on the package rather than the
/// model means every part and every piece of preserved markup is in front of it as XML, and it
/// keeps the four things that must agree in step: parts, relationships, content types and the
/// references inside the XML.
/// </summary>
internal abstract class OpcContentSanitizer(ContentCategories remove)
{
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace OfficeNs = "urn:schemas-microsoft-com:office:office";
    private static readonly XName RelId = OfficeNs + "relid";

    private readonly Dictionary<string, Part> _parts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<Relationship>> _relationships = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<SanitizedItem> _removed = [];
    private readonly List<SanitizedItem> _retained = [];
    private readonly Dictionary<string, ContentCategories> _removedPartCategories = new(StringComparer.OrdinalIgnoreCase);
    private XDocument _contentTypes = new();
    private string _contentTypesEntry = "[Content_Types].xml";

    private sealed class Part(string entryName, byte[] content, DateTimeOffset modified)
    {
        public string EntryName { get; } = entryName;

        public byte[] Content { get; set; } = content;

        public DateTimeOffset Modified { get; } = modified;
    }

    protected sealed record Relationship(string Source, string Id, string Type, string Target, bool IsExternal)
    {
        public ContentCategories Category { get; set; }

        /// <summary>The last segment of the type, which is the same in Transitional and Strict.</summary>
        public string TypeName => Type[(Type.LastIndexOf('/') + 1)..];
    }

    /// <summary>The categories this run removes.</summary>
    protected ContentCategories Selected => remove;

    /// <summary>The element names that mean nothing once they are empty.</summary>
    protected abstract IReadOnlySet<string> PruneWhenEmpty { get; }

    /// <summary>What a macro-enabled main part becomes without its macros, or null to leave it.</summary>
    protected abstract string? PlainContentType(string contentType);

    /// <summary>Makes format-specific changes once references to removed content are gone.</summary>
    protected abstract void ApplyFormatRules();

    /// <summary>Removes an element that refers to a removed relationship.</summary>
    protected virtual void RemoveReference(XElement element, ReferenceScope scope) => scope.Remove(element);

    /// <summary>The category of a relationship, or None when it is ordinary content.</summary>
    protected virtual ContentCategories Classify(Relationship relationship)
    {
        if (relationship.IsExternal)
            return ContentCategories.ExternalLinks;

        return relationship.TypeName switch
        {
            "vbaProject" or "wordVbaData" or "keyMapCustomizations" or "vbaProjectSignature" => ContentCategories.Macros,
            "oleObject" or "package" or "control" or "activeXControlBinary" => ContentCategories.EmbeddedObjects,
            "customXml" or "customXmlProps" => ContentCategories.CustomXml,
            "origin" or "signature" or "certificate" when relationship.Type.Contains("digital-signature", StringComparison.Ordinal) =>
                ContentCategories.Signatures,
            _ => ContentCategories.None,
        };
    }

    public (byte[] Package, SanitizeReport Report) Run(byte[] package)
    {
        Load(package);
        RemoveSelectedRelationships();
        ApplyFormatRules();
        RemoveUnreachableParts();
        RewriteContentTypes();
        return (Write(), new SanitizeReport(_removed, _retained));
    }

    protected void ReportRemoved(ContentCategories category, SanitizedItemKind kind, string location, string detail) =>
        _removed.Add(new SanitizedItem(category, kind, location, detail));

    protected void ReportRetained(ContentCategories category, SanitizedItemKind kind, string location, string detail) =>
        _retained.Add(new SanitizedItem(category, kind, location, detail));

    protected IEnumerable<string> PartPaths => _parts.Keys;

    /// <summary>The parts some relationship of the given type leads to.</summary>
    protected IEnumerable<(string Source, string Target)> TargetsOfType(string typeName) =>
        _relationships.Values
            .SelectMany(static list => list)
            .Where(relationship => !relationship.IsExternal && relationship.TypeName == typeName)
            .Select(static relationship => (relationship.Source, Resolve(relationship)))
            .ToArray();

    protected IReadOnlyList<Relationship> RelationshipsOf(string source) =>
        _relationships.TryGetValue(source, out List<Relationship>? list) ? list : [];

    protected string ReadText(string partPath) => Encoding.UTF8.GetString(_parts[partPath].Content);

    protected static bool IsXmlPart(string partPath) =>
        partPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
        partPath.EndsWith(".vml", StringComparison.OrdinalIgnoreCase);

    protected XDocument ReadXml(string partPath) => Parse(_parts[partPath].Content);

    protected void WriteXml(string partPath, XDocument document) => _parts[partPath].Content = Serialize(document);

    /// <summary>Removes a relationship along with whatever only it was keeping in the package.</summary>
    protected void DropRelationship(Relationship relationship, ContentCategories category)
    {
        if (_relationships.TryGetValue(relationship.Source, out List<Relationship>? list) && list.Remove(relationship))
            MarkRemoved(relationship, category);
    }

    private void MarkRemoved(Relationship relationship, ContentCategories category)
    {
        ReportRemoved(
            category,
            SanitizedItemKind.Relationship,
            relationship.IsExternal ? relationship.Target : Resolve(relationship),
            relationship.TypeName);

        if (!relationship.IsExternal)
            TagDescendants(Resolve(relationship), category);
    }

    /// <summary>Records the category of a part and of everything reachable only through it.</summary>
    private void TagDescendants(string partPath, ContentCategories category)
    {
        var pending = new Stack<string>();
        pending.Push(partPath);
        while (pending.TryPop(out string? current))
        {
            if (!_removedPartCategories.TryAdd(current, category))
                continue;

            foreach (Relationship child in RelationshipsOf(current))
            {
                if (!child.IsExternal)
                    pending.Push(Resolve(child));
            }
        }
    }

    private static string Resolve(Relationship relationship) =>
        OpcPath.Resolve(relationship.Source, relationship.Target);

    private void Load(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using var content = new MemoryStream();
            using (Stream stream = entry.Open())
                stream.CopyTo(content);

            string path = OpcPath.Unescape(OpcPath.ToPartPath(entry.FullName));
            if (path.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            {
                _contentTypes = Parse(content.ToArray());
                _contentTypesEntry = entry.FullName;
            }
            else if (OpcPath.IsRelsPath(path))
            {
                ReadRelationships(OpcPath.GetSourcePart(path), content.ToArray());
            }
            else
            {
                _parts[path] = new Part(entry.FullName, content.ToArray(), entry.LastWriteTime);
            }
        }
    }

    private void ReadRelationships(string source, byte[] content)
    {
        List<Relationship> list = [];
        foreach (XElement element in Parse(content).Root!.Elements(RelationshipNs + "Relationship"))
        {
            list.Add(new Relationship(
                source,
                (string)element.Attribute("Id")!,
                (string)element.Attribute("Type")!,
                (string)element.Attribute("Target")!,
                string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)));
        }

        _relationships[source] = list;
    }

    private void RemoveSelectedRelationships()
    {
        foreach (List<Relationship> list in _relationships.Values)
        {
            foreach (Relationship relationship in list)
                relationship.Category = Classify(relationship);
        }

        foreach ((string source, List<Relationship> list) in _relationships.Select(static pair => (pair.Key, pair.Value)).ToArray())
        {
            var removedIds = new Dictionary<string, ContentCategories>(StringComparer.Ordinal);
            foreach (Relationship relationship in list.ToArray())
            {
                if (relationship.Category == ContentCategories.None)
                    continue;

                if ((relationship.Category & remove) == 0)
                {
                    ReportRetained(
                        relationship.Category,
                        SanitizedItemKind.Relationship,
                        relationship.IsExternal ? relationship.Target : Resolve(relationship),
                        relationship.TypeName);
                    continue;
                }

                removedIds[relationship.Id] = relationship.Category;
                DropRelationship(relationship, relationship.Category);
            }

            if (removedIds.Count > 0 && source != "/" && _parts.ContainsKey(source) && IsXmlPart(source))
                RewriteReferences(source, removedIds);
        }
    }

    /// <summary>
    /// Relationship ids mean something only inside the part that declares them, so only the source
    /// of a removed relationship can hold a reference to it.
    /// </summary>
    private void RewriteReferences(string source, Dictionary<string, ContentCategories> removedIds)
    {
        XDocument document = ReadXml(source);
        var scope = new ReferenceScope(this, source, PruneWhenEmpty);

        foreach (XElement element in document.Descendants().Where(e => ReferencedId(e, removedIds) is not null).ToArray())
        {
            if (element.Document is null)
                continue;

            scope.Category = removedIds[ReferencedId(element, removedIds)!];
            RemoveReference(element, scope);
        }

        DropOrphanedRelationships(source, document, scope);
        WriteXml(source, document);
    }

    private static string? ReferencedId(XElement element, Dictionary<string, ContentCategories> ids)
    {
        foreach (XAttribute attribute in element.Attributes())
        {
            if (IsRelationshipAttribute(attribute) && ids.ContainsKey(attribute.Value))
                return attribute.Value;
        }

        return null;
    }

    private static bool IsRelationshipAttribute(XAttribute attribute) =>
        attribute.Name == RelId ||
        attribute.Name.NamespaceName is
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
            "http://purl.oclc.org/ooxml/officeDocument/relationships";

    /// <summary>
    /// Markup that went away can leave behind what only it used, such as the preview picture of
    /// an embedded object. That picture shows what the object held, so it goes too.
    /// </summary>
    private void DropOrphanedRelationships(string source, XDocument document, ReferenceScope scope)
    {
        if (scope.MentionedIds.Count == 0)
            return;

        HashSet<string> stillUsed = new(StringComparer.Ordinal);
        foreach (XElement element in document.Descendants())
        {
            foreach (XAttribute attribute in element.Attributes())
            {
                if (IsRelationshipAttribute(attribute))
                    stillUsed.Add(attribute.Value);
            }
        }

        foreach (Relationship relationship in RelationshipsOf(source).ToArray())
        {
            if (scope.MentionedIds.TryGetValue(relationship.Id, out ContentCategories category) && !stillUsed.Contains(relationship.Id))
                DropRelationship(relationship, category);
        }
    }

    private void RemoveUnreachableParts()
    {
        HashSet<string> reachable = new(StringComparer.OrdinalIgnoreCase) { "/" };
        var pending = new Stack<string>();
        pending.Push("/");
        while (pending.TryPop(out string? current))
        {
            foreach (Relationship relationship in RelationshipsOf(current))
            {
                if (!relationship.IsExternal && reachable.Add(Resolve(relationship)))
                    pending.Push(Resolve(relationship));
            }
        }

        // Only what this run disconnected goes: a part that was already an orphan in the input is
        // not something the caller asked to have removed.
        foreach ((string path, ContentCategories category) in _removedPartCategories)
        {
            if (reachable.Contains(path) || !_parts.Remove(path))
                continue;

            _relationships.Remove(path);
            ReportRemoved(category, SanitizedItemKind.Part, path, string.Empty);
        }
    }

    private void RewriteContentTypes()
    {
        XElement root = _contentTypes.Root!;
        HashSet<string> extensions = new(StringComparer.OrdinalIgnoreCase) { "rels", "xml" };
        foreach (string path in _parts.Keys)
            extensions.Add(Path.GetExtension(path).TrimStart('.'));

        foreach (XElement entry in root.Elements(ContentTypesNs + "Override").ToArray())
        {
            string partName = OpcPath.Unescape((string)entry.Attribute("PartName")!);
            if (!_parts.ContainsKey(partName))
            {
                entry.Remove();
                continue;
            }

            if ((remove & ContentCategories.Macros) != 0 && PlainContentType((string)entry.Attribute("ContentType")!) is { } plain)
            {
                entry.SetAttributeValue("ContentType", plain);
                ReportRemoved(ContentCategories.Macros, SanitizedItemKind.Markup, partName, "macro-enabled content type");
            }
        }

        foreach (XElement entry in root.Elements(ContentTypesNs + "Default").ToArray())
        {
            if (!extensions.Contains((string)entry.Attribute("Extension")!))
                entry.Remove();
        }
    }

    private byte[] Write()
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, _contentTypesEntry, Serialize(_contentTypes), DateTimeOffset.UnixEpoch);

            foreach (Part part in _parts.Values)
                WriteEntry(archive, part.EntryName, part.Content, part.Modified);

            foreach ((string source, List<Relationship> list) in _relationships)
            {
                if (list.Count > 0 && (source == "/" || _parts.ContainsKey(source)))
                    WriteEntry(archive, OpcPath.ToEntryName(OpcPath.GetRelsPath(source)), SerializeRelationships(list), DateTimeOffset.UnixEpoch);
            }
        }

        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] content, DateTimeOffset modified)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        entry.LastWriteTime = modified.Year < 1980 ? new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero) : modified;
        using Stream stream = entry.Open();
        stream.Write(content);
    }

    private static byte[] SerializeRelationships(List<Relationship> relationships)
    {
        var root = new XElement(RelationshipNs + "Relationships");
        foreach (Relationship relationship in relationships)
        {
            var element = new XElement(
                RelationshipNs + "Relationship",
                new XAttribute("Id", relationship.Id),
                new XAttribute("Type", relationship.Type),
                new XAttribute("Target", relationship.Target));
            if (relationship.IsExternal)
                element.SetAttributeValue("TargetMode", "External");
            root.Add(element);
        }

        return Serialize(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    private static XDocument Parse(byte[] content)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using var reader = XmlReader.Create(new MemoryStream(content), settings);
        return XDocument.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
    }

    private static byte[] Serialize(XDocument document)
    {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false };
        using (var writer = XmlWriter.Create(output, settings))
            document.Save(writer);

        return output.ToArray();
    }

    /// <summary>What one part's rewrite has taken out, and what that markup was pointing at.</summary>
    protected sealed class ReferenceScope(OpcContentSanitizer owner, string source, IReadOnlySet<string> pruneWhenEmpty)
    {
        private const string MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        /// <summary>The category of the reference being handled.</summary>
        public ContentCategories Category { get; set; }

        /// <summary>Ids mentioned inside removed markup, which may now have no other user.</summary>
        public Dictionary<string, ContentCategories> MentionedIds { get; } = new(StringComparer.Ordinal);

        /// <summary>Removes an element, then any container that has lost its only purpose.</summary>
        public void Remove(XElement element)
        {
            foreach (XElement inner in element.DescendantsAndSelf())
            {
                foreach (XAttribute attribute in inner.Attributes())
                {
                    if (IsRelationshipAttribute(attribute))
                        MentionedIds.TryAdd(attribute.Value, Category);
                }
            }

            owner.ReportRemoved(Category, SanitizedItemKind.Markup, source, element.Name.LocalName);

            XElement? parent = element.Parent;
            element.Remove();
            Prune(parent);
        }

        /// <summary>Replaces an element with its children, keeping what it wrapped.</summary>
        public void Unwrap(XElement element)
        {
            owner.ReportRemoved(Category, SanitizedItemKind.Markup, source, element.Name.LocalName);
            element.ReplaceWith(element.Nodes());
        }

        private void Prune(XElement? element)
        {
            while (element?.Parent is { } parent && IsPointless(element))
            {
                element.Remove();
                element = parent;
            }
        }

        private bool IsPointless(XElement element)
        {
            if (element.Name.NamespaceName == MarkupCompatibility)
            {
                return element.Name.LocalName == "AlternateContent"
                    ? !element.Elements(XName.Get("Choice", MarkupCompatibility)).Any()
                    : !element.HasElements;
            }

            return !element.HasElements && pruneWhenEmpty.Contains(element.Name.LocalName);
        }
    }
}
