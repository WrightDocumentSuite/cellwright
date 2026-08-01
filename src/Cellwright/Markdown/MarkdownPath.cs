namespace Cellwright.Markdown;

/// <summary>Maps the media sidecar's relative file-system path to safe Markdown URLs.</summary>
internal static class MarkdownPath
{
    /// <summary>
    /// Normalizes a relative media directory to forward slashes and rejects paths that could
    /// leave the export directory.
    /// </summary>
    public static string NormalizeMediaDirectory(string directory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);

        string normalized = directory.Replace('\\', '/');
        if (Path.IsPathRooted(directory) || normalized.StartsWith('/') || HasDrivePrefix(normalized))
            throw new ArgumentException("The Markdown media directory must be relative.", nameof(directory));

        var segments = new List<string>();
        foreach (string segment in normalized.Split('/'))
        {
            if (segment.Length == 0 || segment == ".")
                continue;
            if (segment == "..")
                throw new ArgumentException("The Markdown media directory cannot contain '..'.", nameof(directory));
            if (segment.Contains('\0'))
                throw new ArgumentException("The Markdown media directory contains an invalid character.", nameof(directory));

            segments.Add(segment);
        }

        if (segments.Count == 0)
            throw new ArgumentException("The Markdown media directory must name a subdirectory.", nameof(directory));

        return string.Join('/', segments);
    }

    /// <summary>Builds a relative URL, encoding each file-system path segment independently.</summary>
    public static string MediaUrl(string directory, string fileName)
    {
        string encodedDirectory = string.Join('/', directory.Split('/').Select(Uri.EscapeDataString));
        return $"{encodedDirectory}/{Uri.EscapeDataString(fileName)}";
    }

    private static bool HasDrivePrefix(string path) =>
        path.Length >= 2 && char.IsAsciiLetter(path[0]) && path[1] == ':';
}
