namespace Cellwright.IO;

/// <summary>
/// Compares part names and media-type extensions the way the packaging conventions require:
/// case-insensitively, but only over ASCII (ECMA-376-2 6.2.2.3 and 7.2.3.5).
/// </summary>
/// <remarks>
/// The distinction is not pedantry. <see cref="StringComparer.OrdinalIgnoreCase"/> folds the
/// whole of Unicode, so it considers the Kelvin sign (U+212A) equal to the letter K, and the
/// Turkish dotted capital I equal to i. Two parts whose names differ only in such a character
/// are distinct parts, and treating them as one silently hands back the contents of the wrong
/// one.
/// </remarks>
internal sealed class OpcNames : IEqualityComparer<string>
{
    /// <summary>The one instance; the comparer has no state.</summary>
    public static OpcNames Instance { get; } = new();

    private OpcNames()
    {
    }

    /// <summary>Whether two part names name the same part.</summary>
    public static bool Same(ReadOnlySpan<char> left, ReadOnlySpan<char> right)
    {
        if (left.Length != right.Length)
            return false;

        for (int i = 0; i < left.Length; i++)
        {
            if (Fold(left[i]) != Fold(right[i]))
                return false;
        }

        return true;
    }

    /// <summary>Whether a name ends with the given ASCII suffix.</summary>
    public static bool EndsWith(string name, string suffix) =>
        name.Length >= suffix.Length && Same(name.AsSpan(name.Length - suffix.Length), suffix);

    bool IEqualityComparer<string>.Equals(string? x, string? y) =>
        x is null || y is null ? ReferenceEquals(x, y) : Same(x, y);

    int IEqualityComparer<string>.GetHashCode(string value)
    {
        var hash = new HashCode();
        foreach (char c in value)
            hash.Add(Fold(c));
        return hash.ToHashCode();
    }

    private static char Fold(char c) => c is >= 'A' and <= 'Z' ? (char)(c + ('a' - 'A')) : c;
}
