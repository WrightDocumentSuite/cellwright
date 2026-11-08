using Cellwright.Tests.Shared;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Finds the legacy workbooks that other spreadsheet libraries keep as test data in sibling
/// repositories. They are the only realistic source of files written by producers nobody here
/// controls, which is exactly what a format reader has to survive.
/// </summary>
internal static class XlsCorpus
{
    private static readonly Lazy<string[]> AllFiles = new(() => Discover("*.xls"));

    private static readonly Lazy<string[]> AllPackages = new(() =>
        [.. Discover("*.xlsx").Concat(Discover("*.xlsm")).OrderBy(static path => path, StringComparer.Ordinal)]);

    /// <summary>Every legacy workbook in the corpus, in a stable order.</summary>
    public static IReadOnlyList<string> Files => AllFiles.Value;

    /// <summary>
    /// Every modern package in the corpus, macro-enabled ones included. The encrypted files are
    /// here rather than among the legacy ones, even though they too are compound files: what is
    /// inside decides.
    /// </summary>
    public static IReadOnlyList<string> Packages => AllPackages.Value;

    /// <summary>Returns up to <paramref name="max"/> workbooks, or an empty list when there is no corpus.</summary>
    public static IReadOnlyList<string> Take(int max) => AllFiles.Value.Take(max).ToArray();

    /// <summary>
    /// Whether an exception means the file itself is unusable rather than the reader being
    /// wrong. Corpora keep deliberately corrupt and encrypted files, and those prove nothing.
    /// </summary>
    /// <remarks>
    /// The package exception is here for the binary format, whose container is a zip: a file
    /// like <c>EmptyZipFile.xlsb</c> fails as a package long before any record is read.
    /// </remarks>
    public static bool IsUnreadable(Exception exception) => exception is
        XlsFormatException or XlsxFormatException or DocumentLoadLimitException or
        InvalidDataException or IOException or NotSupportedException;

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

    /// <summary>Walks up from the test binaries looking for the folder holding the sibling repositories.</summary>
    private static string? FindCorpusRoot()
    {
        return CorpusRoot.Find("ExcelDataReader-3.9.0");
    }
}
