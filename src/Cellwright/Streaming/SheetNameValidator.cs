using System.Buffers;

namespace Cellwright.Streaming;

/// <summary>
/// Enforces Excel's worksheet naming rules.
/// </summary>
internal static class SheetNameValidator
{
    private static readonly SearchValues<char> ForbiddenChars = SearchValues.Create(":\\/?*[]");

    /// <summary>Throws <see cref="ArgumentException"/> when the name violates Excel's rules.</summary>
    public static void Validate(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (name.Length > SheetLimits.MaxSheetNameLength)
            throw new ArgumentException($"Sheet name '{name}' exceeds {SheetLimits.MaxSheetNameLength} characters.", nameof(name));
        if (name.AsSpan().ContainsAny(ForbiddenChars))
            throw new ArgumentException($"Sheet name '{name}' contains one of the forbidden characters : \\ / ? * [ ].", nameof(name));
        if (name.StartsWith('\'') || name.EndsWith('\''))
            throw new ArgumentException($"Sheet name '{name}' cannot start or end with an apostrophe.", nameof(name));
    }
}
