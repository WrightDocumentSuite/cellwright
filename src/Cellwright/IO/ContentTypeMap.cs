using System.Xml;
using Cellwright.Xml;

namespace Cellwright.IO;

/// <summary>
/// The <c>[Content_Types].xml</c> part: extension defaults plus per-part overrides.
/// </summary>
internal sealed class ContentTypeMap
{
    private readonly Dictionary<string, string> _defaults = new(OpcNames.Instance);
    private readonly Dictionary<string, string> _overrides = new(OpcNames.Instance);

    /// <summary>Registers a content type for a file extension (without the dot).</summary>
    public void AddDefault(string extension, string contentType) => _defaults[extension] = contentType;

    /// <summary>Registers a content type override for a specific part.</summary>
    public void AddOverride(string partPath, string contentType) => _overrides[partPath] = contentType;

    /// <summary>Drops the override of a part that is not going into the package.</summary>
    public void RemoveOverride(string partPath) => _overrides.Remove(partPath);

    /// <summary>Extension defaults, keyed by extension without the dot.</summary>
    public IReadOnlyDictionary<string, string> Defaults => _defaults;

    /// <summary>Per-part overrides, keyed by absolute part name.</summary>
    public IReadOnlyDictionary<string, string> Overrides => _overrides;

    /// <summary>Resolves the content type of a part, or <see langword="null"/> when unregistered.</summary>
    /// <remarks>
    /// An override wins over the extension default, and the extension is whatever follows the
    /// last dot of the last segment (ECMA-376-2 7.2.3.5). Searching the whole path for that dot
    /// would give <c>/xl.old/sheet</c> — a part with no extension at all — the extension
    /// <c>old/sheet</c>.
    /// </remarks>
    public string? GetContentType(string partPath)
    {
        if (_overrides.TryGetValue(partPath, out string? overridden))
            return overridden;

        return Extension(partPath) is { } extension && _defaults.TryGetValue(extension, out string? byExtension)
            ? byExtension
            : null;
    }

    /// <summary>The extension of a part name, or <see langword="null"/> when it has none.</summary>
    public static string? Extension(string partPath)
    {
        int segment = partPath.LastIndexOf('/');
        int dot = partPath.LastIndexOf('.');
        return dot > segment ? partPath[(dot + 1)..] : null;
    }

    /// <summary>Parses the content-types XML.</summary>
    public static ContentTypeMap Parse(Stream stream, CancellationToken cancellationToken)
    {
        var map = new ContentTypeMap();
        using var reader = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            if (reader.LocalName == "Default")
            {
                string? extension = reader.GetAttribute("Extension");
                string? contentType = reader.GetAttribute("ContentType");
                if (extension is not null && contentType is not null)
                    map.AddDefault(extension, contentType);
            }
            else if (reader.LocalName == "Override")
            {
                string? partName = reader.GetAttribute("PartName");
                string? contentType = reader.GetAttribute("ContentType");
                if (partName is not null && contentType is not null)
                    map.AddOverride(partName, contentType);
            }
        }

        return map;
    }

    /// <summary>Writes the content-types XML.</summary>
    public void Write(Utf8XmlWriter writer)
    {
        writer.WriteDeclaration();
        writer.WriteRaw("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"u8);
        foreach ((string extension, string contentType) in _defaults)
        {
            writer.WriteRaw("<Default Extension=\""u8);
            writer.WriteAttributeText(extension);
            writer.WriteRaw("\" ContentType=\""u8);
            writer.WriteAttributeText(contentType);
            writer.WriteRaw("\"/>"u8);
        }

        foreach ((string partPath, string contentType) in _overrides)
        {
            writer.WriteRaw("<Override PartName=\""u8);
            writer.WriteAttributeText(partPath);
            writer.WriteRaw("\" ContentType=\""u8);
            writer.WriteAttributeText(contentType);
            writer.WriteRaw("\"/>"u8);
        }

        writer.WriteRaw("</Types>"u8);
    }
}
