namespace Cellwright.IO;

/// <summary>
/// Path arithmetic for OPC part names. Part names are absolute, use forward slashes
/// and are compared case-insensitively.
/// </summary>
internal static class OpcPath
{
    /// <summary>Converts a part name to the zip entry name (strips the leading slash).</summary>
    public static string ToEntryName(string partPath) => partPath.TrimStart('/');

    /// <summary>Converts a zip entry name to an absolute part name.</summary>
    public static string ToPartPath(string entryName) => entryName.StartsWith('/') ? entryName : "/" + entryName;

    /// <summary>Returns the relationships part for a given part, e.g. <c>/xl/workbook.xml</c> to <c>/xl/_rels/workbook.xml.rels</c>. The package root is <c>"/"</c>.</summary>
    public static string GetRelsPath(string partPath)
    {
        if (partPath == "/")
            return "/_rels/.rels";

        int slash = partPath.LastIndexOf('/');
        return $"{partPath[..slash]}/_rels/{partPath[(slash + 1)..]}.rels";
    }

    /// <summary>Returns <see langword="true"/> for a relationships part such as <c>/xl/_rels/workbook.xml.rels</c>.</summary>
    public static bool IsRelsPath(string partPath) =>
        OpcNames.EndsWith(partPath, ".rels") &&
        partPath.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    /// <summary>Inverse of <see cref="GetRelsPath"/>: the part a relationships part describes.</summary>
    public static string GetSourcePart(string relsPath)
    {
        int marker = relsPath.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        string directory = relsPath[..marker];
        string name = relsPath[(marker + "/_rels/".Length)..^".rels".Length];
        return name.Length == 0 ? "/" : $"{directory}/{name}";
    }

    /// <summary>Resolves a relationship target relative to its source part.</summary>
    public static string Resolve(string sourcePartPath, string target)
    {
        if (target.StartsWith('/'))
            return Normalize(target);

        int slash = sourcePartPath.LastIndexOf('/');
        string baseDir = slash <= 0 ? "/" : sourcePartPath[..(slash + 1)];
        return Normalize(baseDir + target);
    }

    /// <summary>
    /// Decodes the URI spelling of a part name without allowing encoded separators or dot
    /// segments to change the package graph.
    /// </summary>
    public static string Unescape(string path)
    {
        if (!path.Contains('%', StringComparison.Ordinal))
            return path;

        string decoded;
        try
        {
            decoded = Uri.UnescapeDataString(path);
        }
        catch (UriFormatException)
        {
            return path;
        }

        if (decoded.Count(static c => c == '/') != path.Count(static c => c == '/') || decoded.Contains('\\'))
            return path;

        foreach (Range range in decoded.AsSpan().Split('/'))
        {
            ReadOnlySpan<char> segment = decoded.AsSpan()[range];
            if (segment is "." or "..")
                return path;
        }

        return decoded;
    }

    /// <summary>
    /// Expresses <paramref name="targetPartPath"/> relative to the part that references it,
    /// which is the form Excel writes and the only one every consumer accepts.
    /// </summary>
    public static string MakeRelative(string sourcePartPath, string targetPartPath)
    {
        string[] source = Segments(sourcePartPath);
        string[] target = Segments(targetPartPath);

        int shared = 0;
        // The last segment of the source is the file name, never part of the shared prefix.
        int sourceDirectoryLength = source.Length - 1;
        while (shared < sourceDirectoryLength && shared < target.Length - 1 &&
               OpcNames.Same(source[shared], target[shared]))
        {
            shared++;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = shared; i < sourceDirectoryLength; i++)
            builder.Append("../");
        for (int i = shared; i < target.Length; i++)
        {
            builder.Append(target[i]);
            if (i < target.Length - 1)
                builder.Append('/');
        }

        return builder.ToString();

        static string[] Segments(string path) => path.TrimStart('/').Split('/');
    }

    /// <summary>Collapses <c>.</c> and <c>..</c> segments.</summary>
    private static string Normalize(string path)
    {
        if (!path.Contains("./", StringComparison.Ordinal))
            return path;

        var segments = new List<string>();
        foreach (Range segmentRange in path.AsSpan().Split('/'))
        {
            ReadOnlySpan<char> segment = path.AsSpan()[segmentRange];
            if (segment.IsEmpty || segment is ".")
                continue;
            if (segment is "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment.ToString());
        }

        return "/" + string.Join('/', segments);
    }
}
