namespace Cellwright.Formulas;

/// <summary>
/// Excel's wildcard matching: <c>*</c> stands for any run of characters, <c>?</c> for one,
/// and <c>~</c> escapes either. Matching is case-insensitive, like the rest of text
/// comparison in formulas.
/// </summary>
internal static class Wildcard
{
    /// <summary>Whether the pattern contains an unescaped wildcard.</summary>
    public static bool ContainsWildcards(ReadOnlySpan<char> pattern)
    {
        for (int i = 0; i < pattern.Length; i++)
        {
            if (pattern[i] == '~')
            {
                i++;
                continue;
            }

            if (pattern[i] is '*' or '?')
                return true;
        }

        return false;
    }

    /// <summary>Whether the whole text matches the pattern.</summary>
    public static bool IsMatch(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
    {
        int textIndex = 0;
        int patternIndex = 0;
        int starText = -1;
        int starPattern = -1;

        while (textIndex < text.Length)
        {
            if (patternIndex < pattern.Length && TryMatchOne(text, textIndex, pattern, ref patternIndex))
            {
                textIndex++;
                continue;
            }

            if (patternIndex < pattern.Length && pattern[patternIndex] == '*')
            {
                starPattern = ++patternIndex;
                starText = textIndex;
                continue;
            }

            if (starPattern < 0)
                return false;

            // Backtrack: let the last star swallow one more character.
            patternIndex = starPattern;
            textIndex = ++starText;
        }

        while (patternIndex < pattern.Length && pattern[patternIndex] == '*')
            patternIndex++;

        return patternIndex == pattern.Length;
    }

    /// <summary>Finds the first position at or after <paramref name="start"/> matching the pattern.</summary>
    public static int IndexOf(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern, int start)
    {
        for (int i = start; i <= text.Length; i++)
        {
            for (int length = text.Length - i; length >= 0; length--)
            {
                if (IsMatch(text.Slice(i, length), pattern))
                    return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Matches one literal or single-character wildcard, advancing the pattern index when it
    /// succeeds.
    /// </summary>
    private static bool TryMatchOne(ReadOnlySpan<char> text, int textIndex, ReadOnlySpan<char> pattern, ref int patternIndex)
    {
        char patternChar = pattern[patternIndex];
        if (patternChar == '?')
        {
            patternIndex++;
            return true;
        }

        if (patternChar == '~' && patternIndex + 1 < pattern.Length)
        {
            if (char.ToUpperInvariant(pattern[patternIndex + 1]) != char.ToUpperInvariant(text[textIndex]))
                return false;
            patternIndex += 2;
            return true;
        }

        if (patternChar == '*')
            return false;

        if (char.ToUpperInvariant(patternChar) != char.ToUpperInvariant(text[textIndex]))
            return false;

        patternIndex++;
        return true;
    }
}
