using Cellwright.Styles;

namespace Cellwright.Pdf.Layout;

/// <summary>Which of the three sections of a header or footer a piece of it belongs to.</summary>
internal enum HeaderSection
{
    /// <summary>Against the left margin.</summary>
    Left,

    /// <summary>In the middle of the page.</summary>
    Center,

    /// <summary>Against the right margin.</summary>
    Right,
}

/// <summary>One stretch of a header or footer that is drawn in one appearance.</summary>
/// <param name="Text">The characters, with every code already resolved.</param>
/// <param name="Font">The font it is drawn in.</param>
internal readonly record struct HeaderPiece(string Text, FontStyle Font);

/// <summary>What one line of a header or footer comes to: three sections of pieces.</summary>
internal sealed class HeaderLine
{
    /// <summary>The pieces of each section, in the order they are drawn.</summary>
    public Dictionary<HeaderSection, List<HeaderPiece>> Sections { get; } = new()
    {
        [HeaderSection.Left] = [],
        [HeaderSection.Center] = [],
        [HeaderSection.Right] = [],
    };

    /// <summary>Whether nothing at all would be drawn.</summary>
    public bool IsEmpty => Sections.Values.All(pieces => pieces.All(piece => piece.Text.Length == 0));
}

/// <summary>
/// Reads the codes a header or footer is written in (ISO/IEC 29500-1 §18.3.1.46).
/// </summary>
/// <remarks>
/// A header is one string carrying its own little markup: <c>&amp;L</c>, <c>&amp;C</c> and
/// <c>&amp;R</c> start a section, <c>&amp;P</c> and <c>&amp;N</c> are the page number and the
/// page count, <c>&amp;A</c> is the sheet's name, and a run of font codes changes how the rest
/// of the section is drawn. What a code cannot resolve to — a picture, a path a file does not
/// have — comes out empty rather than as the code itself.
/// </remarks>
internal static class HeaderFooterParser
{
    /// <summary>What a code stands in for, which only the page it lands on can say.</summary>
    /// <param name="Page">The number this page prints.</param>
    /// <param name="Pages">How many pages the sheet came to.</param>
    /// <param name="Sheet">The sheet's name.</param>
    /// <param name="Now">What the date and time codes print.</param>
    internal readonly record struct Fields(int Page, int Pages, string Sheet, DateTime Now);

    /// <summary>Reads one header or footer.</summary>
    /// <param name="text">The header as the file stores it, codes and all.</param>
    /// <param name="baseFont">The font a section starts in.</param>
    /// <param name="fields">What the codes stand in for.</param>
    public static HeaderLine Parse(string? text, FontStyle baseFont, Fields fields)
    {
        var line = new HeaderLine();
        if (string.IsNullOrEmpty(text))
            return line;

        var section = HeaderSection.Left;
        FontStyle font = baseFont;
        var run = new System.Text.StringBuilder();

        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] != '&' || i + 1 >= text.Length)
            {
                run.Append(text[i]);
                continue;
            }

            char code = text[++i];
            if (Section(code) is { } next)
            {
                Flush(line, section, run, font);
                section = next;
                font = baseFont;
                continue;
            }

            // The text already in hand was written in the font of the moment before the code:
            // in "small&20BIG" only BIG is twenty points, and flushing after the change would
            // print both at that size.
            FontStyle carried = font;
            if (Change(code, text, ref i, ref font, baseFont))
            {
                Flush(line, section, run, carried);
                continue;
            }

            run.Append(Resolve(code, text, ref i, fields));
        }

        Flush(line, section, run, font);
        return line;
    }

    private static HeaderSection? Section(char code) => code switch
    {
        'L' => HeaderSection.Left,
        'C' => HeaderSection.Center,
        'R' => HeaderSection.Right,
        _ => null,
    };

    /// <summary>Ends the run in hand and starts another.</summary>
    private static void Flush(
        HeaderLine line, HeaderSection section, System.Text.StringBuilder run, FontStyle font)
    {
        if (run.Length == 0)
            return;

        line.Sections[section].Add(new HeaderPiece(run.ToString(), font));
        run.Clear();
    }

    /// <summary>
    /// Applies a code that changes how the rest of the section is drawn, and says whether it did.
    /// </summary>
    private static bool Change(char code, string text, ref int i, ref FontStyle font, FontStyle baseFont)
    {
        switch (code)
        {
            case 'B':
                font = font with { Bold = !font.Bold };
                return true;
            case 'I':
                font = font with { Italic = !font.Italic };
                return true;
            case 'U':
                font = font with
                {
                    Underline = font.Underline == UnderlineStyle.None
                        ? UnderlineStyle.Single
                        : UnderlineStyle.None,
                };
                return true;
            case 'S':
                font = font with { Strikethrough = !font.Strikethrough };
                return true;
            case '"':
                font = Face(text, ref i, font);
                return true;
            case >= '0' and <= '9':
                font = font with { Size = Size(text, ref i, code) };
                return true;
            case 'K':
                font = font with { Color = Colour(text, ref i, font.Color) };
                return true;
            case 'X' or 'Y':
                // Superscript and subscript in a header are drawn on the line: the difference
                // is a few points in something that is already the smallest text on the page.
                return true;
            default:
                _ = baseFont;
                return false;
        }
    }

    /// <summary>The face a <c>&amp;"name,style"</c> code names.</summary>
    private static FontStyle Face(string text, ref int i, FontStyle font)
    {
        int close = text.IndexOf('"', i + 1);
        if (close < 0)
            return font;

        string inside = text[(i + 1)..close];
        i = close;

        string[] parts = inside.Split(',');
        string name = parts[0].Trim();
        string style = parts.Length > 1 ? parts[1].Trim() : "";

        return font with
        {
            Name = name is { Length: > 0 } and not "-" ? name : font.Name,
            Bold = style.Contains("Bold", StringComparison.OrdinalIgnoreCase),
            Italic = style.Contains("Italic", StringComparison.OrdinalIgnoreCase),
        };
    }

    /// <summary>The size a run of digits names, in points.</summary>
    private static double Size(string text, ref int i, char first)
    {
        int value = first - '0';
        while (i + 1 < text.Length && char.IsAsciiDigit(text[i + 1]))
            value = (value * 10) + (text[++i] - '0');

        return value > 0 ? value : 11;
    }

    /// <summary>The colour a <c>&amp;K</c> code names, as six hexadecimal digits.</summary>
    private static ExcelColor Colour(string text, ref int i, ExcelColor fallback)
    {
        if (i + 6 >= text.Length || !ExcelColor.TryParseHex(text.AsSpan(i + 1, 6), out ExcelColor color))
            return fallback;

        i += 6;
        return color;
    }

    /// <summary>What a code stands in for, or nothing when it stands in for something undrawable.</summary>
    private static string Resolve(char code, string text, ref int i, Fields fields)
    {
        switch (code)
        {
            case '&':
                return "&";
            case 'P':
                return Page(text, ref i, fields.Page);
            case 'N':
                return fields.Pages.ToString(System.Globalization.CultureInfo.InvariantCulture);
            case 'A':
                return fields.Sheet;
            case 'D':
                return fields.Now.ToString("d", System.Globalization.CultureInfo.InvariantCulture);
            case 'T':
                return fields.Now.ToString("t", System.Globalization.CultureInfo.InvariantCulture);
            case 'F' or 'Z':
                // The file name and its path: a workbook in memory has neither.
                return "";
            case 'G':
                // A picture in a header is not drawn; the diagnostics say so where the header is.
                return "";
            default:
                return "";
        }
    }

    /// <summary>
    /// The page number, which may be offset: <c>&amp;P+2</c> numbers a sheet as if two pages came
    /// before it.
    /// </summary>
    private static string Page(string text, ref int i, int page)
    {
        if (i + 1 < text.Length && text[i + 1] is '+' or '-')
        {
            int sign = text[i + 1] == '+' ? 1 : -1;
            int digits = 0;
            int at = i + 2;

            while (at < text.Length && char.IsAsciiDigit(text[at]))
                digits = (digits * 10) + (text[at++] - '0');

            if (at > i + 2)
            {
                i = at - 1;
                page += sign * digits;
            }
        }

        return page.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}
