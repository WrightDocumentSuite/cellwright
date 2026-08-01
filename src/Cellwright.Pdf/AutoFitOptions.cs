using System.Globalization;

namespace Cellwright.Pdf;

/// <summary>
/// How wide and how tall auto-fitting is allowed to make things.
/// </summary>
/// <example>
/// <code>
/// sheet.AutoFitColumns(new AutoFitOptions { MaxWidth = 40, SampleRows = 500 });
/// </code>
/// </example>
public sealed class AutoFitOptions
{
    /// <summary>The narrowest a fitted column may end up, in characters of the standard font.</summary>
    public double MinWidth { get; init; }

    /// <summary>
    /// The widest a fitted column may end up, in characters. Excel's own ceiling is 255, and a
    /// lower one is usually what a report wants: one long note should not push every other
    /// column off the page.
    /// </summary>
    public double MaxWidth { get; init; } = SheetLimits.MaxColumnWidth;

    /// <summary>The shortest a fitted row may end up, in points.</summary>
    public double MinHeight { get; init; }

    /// <summary>The tallest a fitted row may end up, in points. Excel's ceiling is 409.5.</summary>
    public double MaxHeight { get; init; } = SheetLimits.MaxRowHeight;

    /// <summary>
    /// How many rows to measure, counting from the top of the used range, or <see langword="null"/>
    /// to measure them all. A sheet of a million rows costs a million measurements otherwise, and
    /// the first few hundred usually decide the answer.
    /// </summary>
    public int? SampleRows { get; init; }

    /// <summary>
    /// Whether a hidden row's cells are measured. They are not by default: a row somebody hid is
    /// not one they are sizing the sheet for.
    /// </summary>
    public bool IncludeHiddenRows { get; init; }

    /// <summary>
    /// Whether a column that states no width of its own is left alone. Off by default, so
    /// auto-fitting a fresh sheet does something.
    /// </summary>
    public bool OnlySizedColumns { get; init; }

    /// <summary>
    /// The culture numbers and dates are measured in. It must be the culture they will be
    /// rendered in, or the measurement is of a different string than the one that gets drawn.
    /// </summary>
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>The family used when a style names one that is not installed.</summary>
    public string FallbackFontFamily { get; init; } = "Arial";

    /// <summary>Extra files to resolve font families from, by family name.</summary>
    public IReadOnlyDictionary<string, string> FontFiles { get; init; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>The measuring options as the export options the font machinery speaks.</summary>
    internal PdfExportOptions ToExportOptions()
    {
        var options = new PdfExportOptions
        {
            Culture = Culture,
            FallbackFontFamily = FallbackFontFamily,

            // Subsetting only matters to a document somebody saves, and measuring saves none.
            SubsetFonts = false,
        };

        foreach ((string family, string path) in FontFiles)
            options.FontFiles[family] = path;

        return options;
    }
}
