using System.Net;
using System.Text;

namespace Cellwright.Markdown;

/// <summary>
/// Escaping for text on its way into a Markdown document.
/// </summary>
/// <remarks>
/// CommonMark §2.4 lets a backslash escape any ASCII punctuation, and the escape survives into
/// a table cell (GFM tables, example 200). Only what would otherwise change the structure of the
/// document is escaped: the pipe that would end a cell, the backslash that would swallow the
/// next character, and the newline no table cell can hold. Everything else passes through, which
/// keeps the output readable at the price of letting a cell that reads <c>*note*</c> arrive
/// italic — a compromise, and a deliberate one.
/// </remarks>
internal static class MarkdownText
{
    private const string EscapableAsciiPunctuation = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";

    private static readonly HashSet<string> UnsafeSchemes =
        new(StringComparer.OrdinalIgnoreCase) { "data", "javascript", "vbscript" };

    /// <summary>Appends text as the content of a table cell.</summary>
    public static void AppendCell(StringBuilder builder, string text) => Append(builder, text, linkText: false);

    /// <summary>
    /// Appends text as the label of a link or the alternative text of an image, where the
    /// brackets and parentheses of the surrounding syntax also have to be neutralized.
    /// </summary>
    public static void AppendLinkText(StringBuilder builder, string text) => Append(builder, text, linkText: true);

    /// <summary>Appends a whole link, label and destination both escaped.</summary>
    public static void AppendLink(StringBuilder builder, string text, string destination)
    {
        builder.Append('[');
        AppendLinkText(builder, text);
        builder.Append("](");
        AppendDestination(builder, destination);
        builder.Append(')');
    }

    /// <summary>
    /// Appends the destination of a link. A destination holding a space or a parenthesis is
    /// wrapped in pointy brackets, the form CommonMark provides for exactly that.
    /// </summary>
    public static void AppendDestination(StringBuilder builder, string destination)
    {
        bool pointy = destination.AsSpan().ContainsAny(" \t()");
        if (pointy)
            builder.Append('<');

        foreach (char c in destination)
        {
            switch (c)
            {
                // A destination cannot span lines, and a newline in one is broken markup.
                case '\r' or '\n':
                    continue;
                case '\\' or '<' or '>' or '|':
                    builder.Append('\\');
                    break;
                case '(' or ')' when !pointy:
                    builder.Append('\\');
                    break;
                default:
                    break;
            }

            builder.Append(c);
        }

        if (pointy)
            builder.Append('>');
    }

    /// <summary>
    /// Whether a workbook-provided destination is safe to turn into a clickable link. Markdown
    /// escaping protects the syntax, not the renderer from active URI schemes.
    /// </summary>
    public static bool IsSafeDestination(string destination)
    {
        if (string.IsNullOrWhiteSpace(destination) || destination.Any(char.IsControl))
            return false;

        string decoded = WebUtility.HtmlDecode(destination)
            .Replace("&colon;", ":", StringComparison.OrdinalIgnoreCase)
            .TrimStart();
        if (decoded.Any(char.IsControl))
            return false;

        int colon = decoded.IndexOf(':');
        if (colon <= 0 || !IsScheme(decoded.AsSpan(0, colon)))
            return true;

        return !UnsafeSchemes.Contains(decoded[..colon]);
    }

    /// <summary>
    /// Appends text as a heading. A heading is one line by definition, so the line breaks a
    /// sheet name or a document title may carry become spaces.
    /// </summary>
    public static void AppendHeading(StringBuilder builder, string text)
    {
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c is '\r' or '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
                builder.Append(' ');
            }
            else
            {
                if (EscapableAsciiPunctuation.Contains(c))
                    builder.Append('\\');
                builder.Append(c);
            }
        }
    }

    private static void Append(StringBuilder builder, string text, bool linkText)
    {
        int leadingEnd = 0;
        while (leadingEnd < text.Length && IsTableEdgeWhitespace(text[leadingEnd]))
            leadingEnd++;

        int trailingStart = text.Length;
        while (trailingStart > leadingEnd && IsTableEdgeWhitespace(text[trailingStart - 1]))
            trailingStart--;

        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if ((i < leadingEnd || i >= trailingStart) && IsTableEdgeWhitespace(c))
            {
                builder.Append(c == ' ' ? "&nbsp;" : "&#9;");
                continue;
            }

            switch (c)
            {
                case '\r' or '\n':
                    // A cell holding several lines becomes one line with breaks in it: a table
                    // cell cannot contain a block, and CRLF must not turn into two of them.
                    if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                        i++;
                    builder.Append("<br>");
                    break;
                case '\\' or '|':
                    builder.Append('\\').Append(c);
                    break;
                case '[' or ']' or '(' or ')' when linkText:
                    builder.Append('\\').Append(c);
                    break;
                default:
                    builder.Append(c);
                    break;
            }
        }
    }

    private static bool IsScheme(ReadOnlySpan<char> candidate)
    {
        if (!char.IsAsciiLetter(candidate[0]))
            return false;

        foreach (char c in candidate[1..])
        {
            if (!char.IsAsciiLetterOrDigit(c) && c is not '+' and not '-' and not '.')
                return false;
        }

        return true;
    }

    private static bool IsTableEdgeWhitespace(char c) => c is ' ' or '\t';
}
