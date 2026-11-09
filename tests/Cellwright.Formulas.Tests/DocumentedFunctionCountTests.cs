using System.Text.RegularExpressions;
using Cellwright.Formulas.Functions;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// The number of supported functions is quoted in several places that ship separately: the
/// repository README, the NuGet README and the package description. They drifted apart once, so
/// each is held to the registry, which is the only place the number is actually decided.
/// </summary>
public sealed class DocumentedFunctionCountTests
{
    private static string RepositoryRoot()
    {
        for (DirectoryInfo? directory = new(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Cellwright.slnx")))
                return directory.FullName;
        }

        throw new DirectoryNotFoundException("The tests are not running from inside the repository.");
    }

    [Theory]
    [InlineData("README.md", @"`Cellwright\.Formulas`\): (\d+) functions")]
    [InlineData("README.nuget.md", @"Calculation engine: (\d+) functions")]
    [InlineData("src/Cellwright.Formulas/Cellwright.Formulas.csproj", @"(\d+) worksheet functions")]
    [InlineData("docs/formulas.md", @"## Supported functions \((\d+)\)")]
    public void TheDocumentedCountMatchesTheRegistry(string relativePath, string pattern)
    {
        string text = File.ReadAllText(Path.Combine(RepositoryRoot(), relativePath));

        Match match = Regex.Match(text, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        Assert.True(match.Success, $"{relativePath} no longer states the function count in the expected form.");
        Assert.Equal(FunctionRegistry.Names.Count, int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
    }
}
