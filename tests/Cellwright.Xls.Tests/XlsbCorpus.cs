using Cellwright.Tests.Shared;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Finds the binary workbooks that other spreadsheet libraries keep as test data in sibling
/// repositories, the same way <see cref="XlsCorpus"/> finds the legacy ones.
/// </summary>
/// <remarks>
/// A dozen of them are encrypted, which is not a defect in the file: the reader opens those
/// through the same path as an encrypted xlsx, so they are kept in and skipped only where a
/// test has no password to offer.
/// </remarks>
internal static class XlsbCorpus
{
    private static readonly Lazy<string[]> AllFiles = new(() => Discover("*.xlsb"));

    private static readonly Lazy<(string Binary, string Twin)[]> AllTwins = new(DiscoverTwins);

    /// <summary>Every binary workbook in the corpus, in a stable order.</summary>
    public static IReadOnlyList<string> Files => AllFiles.Value;

    /// <summary>
    /// Binary workbooks that the corpus also keeps in another format, paired with that other
    /// copy. The same fixture saved twice is the only oracle this format has for anything the
    /// value comparison does not reach.
    /// </summary>
    public static IReadOnlyList<(string Binary, string Twin)> Twins => AllTwins.Value;

    private static string[] Discover(string pattern)
    {
        string? root = FindCorpusRoot();
        if (root is null)
            return [];

        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
        return [.. Directory.EnumerateFiles(root, pattern, options)
            .Where(CorpusRoot.Includes)
            .OrderBy(static path => path, StringComparer.Ordinal)];
    }

    /// <summary>
    /// Pairs each binary workbook with a copy of the same fixture in another format, matched by
    /// file name. A legacy twin is preferred over a package one: the legacy reader has been
    /// checked against two other implementations over a far larger corpus, so it is the better
    /// witness of what the file is supposed to say.
    /// </summary>
    private static (string Binary, string Twin)[] DiscoverTwins()
    {
        string? root = FindCorpusRoot();
        if (root is null)
            return [];

        // Generic fixture names such as "Simple" and "hyperlink" occur independently in
        // several upstream repositories. Pairing them globally can compare unrelated books
        // just because their base names happen to match, so the archive/repository directory
        // is part of the identity.
        var candidates = new Dictionary<string, (string? Legacy, string? Package)>(StringComparer.OrdinalIgnoreCase);
        foreach (string path in Discover("*.xls").Concat(Discover("*.xlsx")).Concat(Discover("*.xlsm")))
        {
            // A Strict package states the same content under different namespaces, and the
            // writer of the fixture was not the same tool; it is a copy of the file, not of
            // the workbook, so it makes a poor witness.
            if (path.Contains("strict", StringComparison.OrdinalIgnoreCase))
                continue;

            string name = Path.GetFileNameWithoutExtension(path);
            string key = CorpusKey(root, path, name);
            candidates.TryGetValue(key, out (string? Legacy, string? Package) found);
            candidates[key] = path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
                ? (found.Legacy ?? path, found.Package)
                : (found.Legacy, found.Package ?? path);
        }

        var pairs = new List<(string, string)>();
        foreach (string binary in AllFiles.Value)
        {
            if (IsKnownDifferentFixture(root, binary))
                continue;

            string key = CorpusKey(root, binary, Path.GetFileNameWithoutExtension(binary));
            if (candidates.TryGetValue(key, out (string? Legacy, string? Package) twin) &&
                (twin.Legacy ?? twin.Package) is { } chosen)
            {
                pairs.Add((binary, chosen));
            }
        }

        return [.. pairs];
    }

    private static string CorpusKey(string root, string path, string name)
    {
        return string.Concat(CorpusSource(root, path), "\0", name);
    }

    private static string CorpusSource(string root, string path)
    {
        ReadOnlySpan<char> relative = Path.GetRelativePath(root, path);
        int separator = relative.IndexOfAny(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return separator < 0 ? "." : relative[..separator].ToString();
    }

    private static bool IsKnownDifferentFixture(string root, string binary)
    {
        // NPOI's comments.xlsb has an A2 comment which its comments.xls and comments.xlsx do
        // not contain. The raw files therefore are not two encodings of the same workbook and
        // cannot witness either reader, even though their base names match.
        return Path.GetFileNameWithoutExtension(binary).Equals("comments", StringComparison.OrdinalIgnoreCase) &&
               CorpusSource(root, binary).StartsWith("npoi-", StringComparison.OrdinalIgnoreCase);
    }

    private static string? FindCorpusRoot()
    {
        return CorpusRoot.Find("ExcelDataReader-3.9.0");
    }
}
