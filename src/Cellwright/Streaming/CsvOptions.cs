using System.Globalization;
using System.Text;

namespace Cellwright;

/// <summary>
/// Options controlling CSV and TSV reading and writing.
/// </summary>
public sealed record CsvOptions
{
    private static readonly UTF8Encoding Utf8WithoutBom = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Standard comma-separated options (<c>,</c> delimiter, <c>"</c> quote, invariant culture).</summary>
    public static CsvOptions Default { get; } = new();

    /// <summary>Tab-separated options (<c>\t</c> delimiter, <c>"</c> quote, invariant culture).</summary>
    public static CsvOptions TabSeparated { get; } = new() { Delimiter = '\t' };

    /// <summary>Semicolon-separated options (<c>;</c> delimiter, <c>"</c> quote, invariant culture).</summary>
    public static CsvOptions SemicolonSeparated { get; } = new() { Delimiter = ';' };

    /// <summary>Field delimiter character. Defaults to <c>,</c>.</summary>
    public char Delimiter { get; init; } = ',';

    /// <summary>Quote character used to enclose fields containing delimiters, quotes, or line breaks. Defaults to <c>"</c>.</summary>
    public char Quote { get; init; } = '"';

    /// <summary>Text encoding. Defaults to UTF-8 (with BOM auto-detection on read and no BOM on write).</summary>
    public Encoding Encoding { get; init; } = Utf8WithoutBom;

    /// <summary>Culture used when parsing or formatting numbers and dates. Defaults to <see cref="CultureInfo.InvariantCulture"/>.</summary>
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Whether unquoted fields should be parsed into numbers, booleans, dates, durations, or
    /// Excel errors when reading. Quoted fields are always preserved as text unless
    /// <see cref="InferQuotedTypes"/> is also enabled. Defaults to <see langword="true"/>.
    /// </summary>
    public bool InferTypes { get; init; } = true;

    /// <summary>
    /// Whether quoted fields (e.g. <c>"42"</c>) should also undergo type inference when
    /// <see cref="InferTypes"/> is <see langword="true"/>. Defaults to <see langword="false"/>
    /// so quoted leading-zero codes such as <c>"00123"</c> remain text.
    /// </summary>
    public bool InferQuotedTypes { get; init; }

    /// <summary>
    /// Whether worksheet export formats values through each cell's number format
    /// instead of raw values. Defaults to <see langword="false"/>.
    /// </summary>
    public bool UseFormattedValues { get; init; }

    /// <summary>Line terminator written between records. Defaults to <c>\r\n</c> (RFC 4180).</summary>
    public string NewLine { get; init; } = "\r\n";

    /// <summary>Worksheet name used when loading a CSV into a <see cref="Workbook"/>. Defaults to <c>"Sheet1"</c>.</summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>Whether to leave the underlying stream open after reading or writing. Defaults to <see langword="false"/>.</summary>
    public bool LeaveOpen { get; init; }

    /// <summary>
    /// Whether a malformed quoted field is an error: text after a field's closing quote, or a
    /// quote that is never closed. When <see langword="false"/>, the default, such a field is
    /// recovered the way Excel recovers it, by keeping the text after the closing quote as part
    /// of the field. When <see langword="true"/>, reading raises <see cref="InvalidDataException"/>.
    /// </summary>
    public bool StrictQuotes { get; init; }

    /// <summary>Resource limits enforced when reading CSV/TSV inputs. Defaults to <see cref="DocumentLoadBudget.Default"/>.</summary>
    public DocumentLoadBudget LoadBudget { get; init; } = DocumentLoadBudget.Default;

    internal void Validate()
    {
        if (Delimiter == Quote)
            throw new ArgumentException("Delimiter and Quote cannot be the same character.");
        if (Delimiter is '\r' or '\n' || Quote is '\r' or '\n')
            throw new ArgumentException("Delimiter and Quote cannot be line break characters.");
        ArgumentException.ThrowIfNullOrEmpty(NewLine);
        ArgumentException.ThrowIfNullOrWhiteSpace(SheetName);
        ArgumentNullException.ThrowIfNull(LoadBudget);
        LoadBudget.Validate();
    }
}
