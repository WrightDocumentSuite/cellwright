using Cellwright.Styles;
using Inkwright;
using Inkwright.Fonts;

namespace Cellwright.Pdf.Fonts;

/// <summary>
/// Turns the font a cell asks for into a font the PDF can draw with.
/// </summary>
/// <remarks>
/// <para>
/// A workbook names a family and leaves it to the machine to find a file. That works on the
/// author's desktop and fails on a build server, so the lookup here is a chain rather than a
/// single try: the caller's own files, then the machine's fonts, then a table of faces that are
/// metrically or visually close, and finally one of the fourteen fonts every reader carries.
/// Whenever the chain falls past the family the workbook asked for, the substitution is recorded
/// in <see cref="PdfExportDiagnostics"/> instead of being silently accepted.
/// </para>
/// <para>
/// A face is embedded once per document however many cells use it, because
/// <see cref="PdfFontCollection"/> already interns by file path.
/// </para>
/// </remarks>
internal sealed class FontMap
{
    /// <summary>
    /// Faces that stand in for a family that is not installed, in descending order of fidelity.
    /// The first entries are the metric-compatible clones shipped with Linux distributions, which
    /// keep a column the width the author sized it to.
    /// </summary>
    private static readonly Dictionary<string, string[]> Substitutes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = ["Carlito", "Segoe UI", "Arial", "Liberation Sans", "DejaVu Sans"],
        ["Calibri Light"] = ["Carlito", "Segoe UI", "Calibri", "Arial", "Liberation Sans"],
        ["Cambria"] = ["Caladea", "Georgia", "Times New Roman", "Liberation Serif", "DejaVu Serif"],
        ["Times New Roman"] = ["Liberation Serif", "Tinos", "Georgia", "DejaVu Serif"],
        ["Arial"] = ["Liberation Sans", "Arimo", "Helvetica", "DejaVu Sans"],
        ["Helvetica"] = ["Arial", "Liberation Sans", "Arimo", "DejaVu Sans"],
        ["Courier New"] = ["Liberation Mono", "Cousine", "DejaVu Sans Mono"],
        ["Segoe UI"] = ["Selawik", "Calibri", "Carlito", "Arial", "Liberation Sans"],
        ["Verdana"] = ["DejaVu Sans", "Bitstream Vera Sans", "Arial", "Liberation Sans"],
        ["Tahoma"] = ["Verdana", "DejaVu Sans", "Arial", "Liberation Sans"],
        ["Georgia"] = ["Gelasio", "Times New Roman", "Liberation Serif", "DejaVu Serif"],
        ["Consolas"] = ["Inconsolata", "Courier New", "Liberation Mono", "DejaVu Sans Mono"],
        ["Trebuchet MS"] = ["Fira Sans", "Verdana", "Arial", "Liberation Sans"],
        ["Wingdings"] = ["Wingdings 2", "Wingdings 3", "Symbol"],
        ["Webdings"] = ["Wingdings", "Symbol"],
        ["Symbol"] = ["OpenSymbol", "Standard Symbols PS"],
    };

    /// <summary>Families whose glyphs are monospaced, used to pick the built-in fallback.</summary>
    private static readonly string[] MonospaceHints = ["courier", "mono", "consol", "menlo", "monaco"];

    /// <summary>Families with serifs, used to pick the built-in fallback.</summary>
    private static readonly string[] SerifHints =
        ["times", "serif", "georgia", "cambria", "garamond", "book", "palatino", "roman", "constantia"];

    private readonly PdfDocument _pdf;
    private readonly PdfExportOptions _options;
    private readonly PdfExportDiagnostics _diagnostics;
    private readonly Dictionary<(string Family, bool Bold, bool Italic), PdfFont> _cache = [];

    internal FontMap(PdfDocument pdf, PdfExportOptions options, PdfExportDiagnostics diagnostics)
    {
        _pdf = pdf;
        _options = options;
        _diagnostics = diagnostics;
    }

    /// <summary>The font a cell's character formatting draws with.</summary>
    /// <param name="font">The font a style names.</param>
    public PdfFont Resolve(FontStyle font)
    {
        ArgumentNullException.ThrowIfNull(font);
        return Resolve(font.Name, font.Bold, font.Italic);
    }

    /// <summary>The font a family and style draw with, after substitution.</summary>
    /// <param name="family">The family the workbook asks for.</param>
    /// <param name="bold">Whether a bold face is wanted.</param>
    /// <param name="italic">Whether an italic face is wanted.</param>
    public PdfFont Resolve(string? family, bool bold, bool italic)
    {
        string wanted = string.IsNullOrWhiteSpace(family) ? _options.FallbackFontFamily : family.Trim();

        if (_cache.TryGetValue((wanted, bold, italic), out PdfFont? cached))
            return cached;

        PdfFont font = Locate(wanted, bold, italic);
        _cache[(wanted, bold, italic)] = font;
        return font;
    }

    private PdfFont Locate(string family, bool bold, bool italic)
    {
        if (Embed(family, bold, italic) is { } exact)
            return exact;

        if (Substitutes.TryGetValue(family, out string[]? chain))
        {
            foreach (string candidate in chain)
            {
                if (Embed(candidate, bold, italic) is not { } substitute)
                    continue;

                _diagnostics.Add(
                    PdfExportWarningKind.FontSubstituted,
                    $"'{family}' is not available; '{candidate}' was drawn instead.",
                    family);
                return substitute;
            }
        }

        if (!family.Equals(_options.FallbackFontFamily, StringComparison.OrdinalIgnoreCase) &&
            Embed(_options.FallbackFontFamily, bold, italic) is { } fallback)
        {
            _diagnostics.Add(
                PdfExportWarningKind.FontSubstituted,
                $"'{family}' is not available; the fallback '{_options.FallbackFontFamily}' was drawn instead.",
                family);
            return fallback;
        }

        Standard14Font builtIn = BuiltIn(family, bold, italic);
        _diagnostics.Add(
            PdfExportWarningKind.FontSubstituted,
            $"'{family}' is not available and no substitute is installed; the built-in '{builtIn.Name}' was drawn instead.",
            family);

        return _pdf.Fonts.Standard(builtIn);
    }

    /// <summary>Embeds a face from the caller's own files or from the machine, or answers null.</summary>
    private PdfFont? Embed(string family, bool bold, bool italic)
    {
        if (FindFile(family, bold, italic) is { } path)
            return _pdf.Fonts.Embed(path, _options.SubsetFonts);

        string? system = FindSystemFont(family, bold, italic);
        return system is null ? null : _pdf.Fonts.Embed(system, _options.SubsetFonts);
    }

    /// <summary>
    /// Resolves installed Office faces while remaining compatible with Inkwright releases that
    /// indexed only font filenames. Windows uses abbreviated files such as <c>times.ttf</c> and
    /// <c>cour.ttf</c>, which cannot otherwise be found from their workbook family names.
    /// </summary>
    private static string? FindSystemFont(string family, bool bold, bool italic)
    {
        string? path = SystemFonts.Find(family, bold, italic);
        if (path is not null || !OperatingSystem.IsWindows())
            return path;

        string compact = new(family.Where(char.IsAsciiLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        string? stem = compact switch
        {
            _ when compact.StartsWith("timesnewroman", StringComparison.Ordinal) || compact == "timesroman" => "times",
            _ when compact.StartsWith("couriernew", StringComparison.Ordinal) || compact == "courier" => "cour",
            _ when compact.StartsWith("arial", StringComparison.Ordinal) => "arial",
            _ => null,
        };
        if (stem is null)
            return null;

        string suffix = (bold, italic) switch
        {
            (true, true) => "bi",
            (true, false) => "bd",
            (false, true) => "i",
            _ => string.Empty,
        };
        string candidate = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", stem + suffix + ".ttf");
        return File.Exists(candidate) ? candidate : null;
    }

    /// <summary>
    /// The caller's own file for a face. The style is looked for as a suffix on the family name
    /// first, so a caller can register <c>Calibri</c> and <c>Calibri Bold</c> separately, and the
    /// plain family is accepted as a last resort.
    /// </summary>
    private string? FindFile(string family, bool bold, bool italic)
    {
        if (_options.FontFiles.Count == 0)
            return null;

        foreach (string key in StyleKeys(family, bold, italic))
        {
            if (_options.FontFiles.TryGetValue(key, out string? path) && File.Exists(path))
                return path;
        }

        return null;
    }

    private static IEnumerable<string> StyleKeys(string family, bool bold, bool italic)
    {
        if (bold && italic)
        {
            yield return family + " Bold Italic";
            yield return family + " BoldItalic";
        }

        if (bold)
            yield return family + " Bold";

        if (italic)
            yield return family + " Italic";

        yield return family;
    }

    /// <summary>
    /// The one of the fourteen built-in fonts closest to a family, chosen by what the name says
    /// about it. A face nobody has is still better drawn in the right shape than not at all.
    /// </summary>
    private static Standard14Font BuiltIn(string family, bool bold, bool italic)
    {
        string lowered = family.ToLowerInvariant();

        if (lowered.Contains("wingding", StringComparison.Ordinal) ||
            lowered.Contains("dingbat", StringComparison.Ordinal) ||
            lowered.Contains("webding", StringComparison.Ordinal))
        {
            return Standard14Font.ZapfDingbats;
        }

        if (lowered.Equals("symbol", StringComparison.Ordinal))
            return Standard14Font.Symbol;

        if (MonospaceHints.Any(hint => lowered.Contains(hint, StringComparison.Ordinal)))
        {
            return (bold, italic) switch
            {
                (true, true) => Standard14Font.Find("Courier-BoldOblique")!,
                (true, false) => Standard14Font.CourierBold,
                (false, true) => Standard14Font.Find("Courier-Oblique")!,
                _ => Standard14Font.Courier,
            };
        }

        if (SerifHints.Any(hint => lowered.Contains(hint, StringComparison.Ordinal)))
        {
            return (bold, italic) switch
            {
                (true, true) => Standard14Font.TimesBoldItalic,
                (true, false) => Standard14Font.TimesBold,
                (false, true) => Standard14Font.TimesItalic,
                _ => Standard14Font.TimesRoman,
            };
        }

        return (bold, italic) switch
        {
            (true, true) => Standard14Font.HelveticaBoldOblique,
            (true, false) => Standard14Font.HelveticaBold,
            (false, true) => Standard14Font.HelveticaOblique,
            _ => Standard14Font.Helvetica,
        };
    }
}
