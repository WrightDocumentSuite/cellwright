namespace Cellwright.Tests;

/// <summary>
/// Locates the sample workbooks that ship with the neighbouring spreadsheet libraries. They
/// are the only source of the messy reality a synthetic fixture never reproduces: charts,
/// pivot caches, VML, macros, legacy producers and outright damaged files.
/// </summary>
internal static class SampleCorpus
{
    // Macro-enabled workbooks are the same format with another extension, and they carry the
    // one part with the most to lose, so the fidelity corpus covers both.
    private static readonly Lazy<string[]> AllFiles = new(() =>
        [.. Discover("*.xlsx").Concat(Discover("*.xlsm")).OrderBy(static path => path, StringComparer.Ordinal)]);

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string[]> ByExtension = new();

    /// <summary>
    /// Whether a file simply cannot be read. Other suites keep corrupt, encrypted and
    /// hand-truncated workbooks around on purpose, and those are not our fidelity signal.
    /// </summary>
    public static bool IsUnreadable(Exception exception) => exception is
        XlsxFormatException or InvalidDataException or IOException or
        NotSupportedException or InvalidOperationException or
        System.Xml.XmlException or ArgumentException;

    /// <summary>Returns up to <paramref name="max"/> sample workbooks, or an empty list.</summary>
    public static IReadOnlyList<string> Files(int max) => AllFiles.Value.Take(max).ToArray();

    /// <summary>
    /// Finds one corpus file by name, whatever its extension. Files of formats this library
    /// does not read live in the same folders, and the tests that prove they are turned away
    /// politely need to reach them.
    /// </summary>
    public static string? Find(string fileName)
    {
        string[] candidates = ByExtension.GetOrAdd("*" + Path.GetExtension(fileName), Discover);
        return candidates.FirstOrDefault(path =>
            Path.GetFileName(path).Equals(fileName, StringComparison.OrdinalIgnoreCase));
    }

    private static string[] Discover(string pattern)
    {
        string? root = FindCorpusRoot();
        if (root is null)
            return [];

        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
        return [.. Directory.EnumerateFiles(root, pattern, options)
            .Where(static path => !path.Contains("Cellwright", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static path => path, StringComparer.Ordinal)];
    }

    /// <summary>Walks up from the test binaries looking for the folder holding the sibling repositories.</summary>
    private static string? FindCorpusRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "ClosedXML-0.105.0")))
                return directory.FullName;
            directory = directory.Parent;
        }

        return null;
    }
}
