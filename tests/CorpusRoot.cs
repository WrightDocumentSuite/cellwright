namespace Cellwright.Tests.Shared;

/// <summary>Resolves the optional, explicitly verified integration corpus.</summary>
internal static class CorpusRoot
{
    private const string EnvironmentVariable = "CELLWRIGHT_CORPUS_ROOT";

    /// <summary>
    /// Uses the configured corpus exclusively, or falls back to the historical sibling-repository
    /// discovery used by local development.
    /// </summary>
    public static string? Find(string sentinelDirectory)
    {
        if (Environment.GetEnvironmentVariable(EnvironmentVariable) is { Length: > 0 } configured)
        {
            string root = Path.GetFullPath(configured);
            if (!Directory.Exists(root))
                throw new DirectoryNotFoundException($"Configured corpus root does not exist: {root}");
            if (!Directory.Exists(Path.Combine(root, sentinelDirectory)))
            {
                throw new DirectoryNotFoundException(
                    $"Configured corpus root '{root}' does not contain '{sentinelDirectory}'.");
            }

            return root;
        }

        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, sentinelDirectory)))
                return directory.FullName;
            directory = directory.Parent;
        }

        return null;
    }

    /// <summary>
    /// The fallback root can also contain this repository, whose own fixtures are not an external
    /// oracle. An explicit root has already been assembled for the integration run and is exact.
    /// </summary>
    public static bool Includes(string path) =>
        Environment.GetEnvironmentVariable(EnvironmentVariable) is { Length: > 0 } ||
        !path.Contains("Cellwright", StringComparison.OrdinalIgnoreCase);
}
