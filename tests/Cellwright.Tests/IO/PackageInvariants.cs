using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Cellwright.IO;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// What every package has to get right whatever it contains. Comparing text after a round trip
/// cannot see any of this: a duplicate entry or a dangling relationship leaves the words intact
/// and still makes the file unopenable.
/// </summary>
internal static class PackageInvariants
{
    /// <summary>Everything structurally wrong with a package, one line per problem.</summary>
    /// <param name="package">The package; left open and at an unspecified position.</param>
    public static HashSet<string> Problems(Stream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);

        var problems = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = OpcPath.Unescape(OpcPath.ToPartPath(entry.FullName));
            if (!names.Add(name))
                problems.Add($"{name} appears more than once");
        }

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = OpcPath.Unescape(OpcPath.ToPartPath(entry.FullName));
            if (OpcPath.IsRelsPath(name))
                CheckRelationships(entry, name, names, problems);
        }

        CheckContentTypes(archive, names, problems);
        return problems;
    }

    /// <summary>Fails if a save introduced a problem the source package did not already have.</summary>
    /// <param name="sourcePath">The file that was loaded.</param>
    /// <param name="saved">What it was saved as; rewound afterwards.</param>
    public static void AssertNoNewProblems(string sourcePath, Stream saved)
    {
        HashSet<string> known;
        using (Stream source = File.OpenRead(sourcePath))
            known = Problems(source);

        string[] introduced = [.. Problems(saved).Except(known, StringComparer.Ordinal).Order(StringComparer.Ordinal)];
        saved.Position = 0;
        Assert.True(introduced.Length == 0, $"{Path.GetFileName(sourcePath)}: {string.Join("; ", introduced)}");
    }

    private static void CheckRelationships(ZipArchiveEntry entry, string name, HashSet<string> names, HashSet<string> problems)
    {
        XDocument relationships;
        try
        {
            using Stream stream = entry.Open();
            relationships = XDocument.Load(stream);
        }
        catch (XmlException)
        {
            problems.Add($"{name} is not well-formed");
            return;
        }

        string source = OpcPath.GetSourcePart(name);
        foreach (XElement relationship in relationships.Root!.Elements())
        {
            if ((string?)relationship.Attribute("TargetMode") == "External" || (string?)relationship.Attribute("Target") is not { } target)
                continue;

            string resolved = OpcPath.Resolve(source, target);
            if (!names.Contains(resolved))
                problems.Add($"{name} points at {resolved}, which is not in the package");
        }
    }

    private static void CheckContentTypes(ZipArchive archive, HashSet<string> names, HashSet<string> problems)
    {
        if (archive.GetEntry("[Content_Types].xml") is not { } entry)
        {
            problems.Add("the package has no content types");
            return;
        }

        XDocument types;
        try
        {
            using Stream stream = entry.Open();
            types = XDocument.Load(stream);
        }
        catch (XmlException)
        {
            problems.Add("[Content_Types].xml is not well-formed");
            return;
        }

        var overrides = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement type in types.Root!.Elements())
        {
            if (type.Name.LocalName == "Default" && (string?)type.Attribute("Extension") is { } extension)
            {
                extensions.Add(extension);
            }
            else if (type.Name.LocalName == "Override" && (string?)type.Attribute("PartName") is { } part)
            {
                string partName = OpcPath.Unescape(part);
                overrides.Add(partName);
                if (!names.Contains(partName))
                    problems.Add($"a content type names {partName}, which is not in the package");
            }
        }

        foreach (string name in names)
        {
            if (!name.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase) &&
                !overrides.Contains(name) &&
                !extensions.Contains(Path.GetExtension(name).TrimStart('.')))
            {
                problems.Add($"{name} has no content type");
            }
        }
    }
}
