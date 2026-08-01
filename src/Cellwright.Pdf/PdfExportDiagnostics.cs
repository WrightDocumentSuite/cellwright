using System.Collections;

namespace Cellwright.Pdf;

/// <summary>What kind of compromise the exporter had to make.</summary>
public enum PdfExportWarningKind
{
    /// <summary>A font the workbook names was not available and another face was drawn instead.</summary>
    FontSubstituted,

    /// <summary>An image could not be embedded, usually because of its format.</summary>
    ImageSkipped,

    /// <summary>Content the renderer does not draw was left out.</summary>
    ContentSkipped,

    /// <summary>A layout constraint could not be honoured and something was approximated.</summary>
    LayoutApproximated,
}

/// <summary>One compromise made while rendering, with enough context to act on it.</summary>
/// <param name="Kind">What kind of compromise it was.</param>
/// <param name="Message">What happened, in one sentence.</param>
/// <param name="Subject">The sheet, font family or feature involved.</param>
public readonly record struct PdfExportWarning(PdfExportWarningKind Kind, string Message, string? Subject = null)
{
    /// <inheritdoc />
    public override string ToString() => Subject is null ? Message : $"{Message} ({Subject})";
}

/// <summary>
/// Everything the exporter had to work around. A render never throws for content it cannot
/// draw: it draws what it can and says here what it left out.
/// </summary>
public sealed class PdfExportDiagnostics : IReadOnlyList<PdfExportWarning>
{
    private readonly List<PdfExportWarning> _warnings = [];
    private readonly HashSet<(PdfExportWarningKind Kind, string? Subject)> _seen = [];

    /// <inheritdoc />
    public int Count => _warnings.Count;

    /// <inheritdoc />
    public PdfExportWarning this[int index] => _warnings[index];

    /// <summary>Whether anything at all had to be worked around.</summary>
    public bool IsEmpty => _warnings.Count == 0;

    /// <summary>The families that were drawn with a different face, in the order they were met.</summary>
    public IEnumerable<string> SubstitutedFonts => _warnings
        .Where(static warning => warning.Kind == PdfExportWarningKind.FontSubstituted && warning.Subject is not null)
        .Select(static warning => warning.Subject!);

    /// <inheritdoc />
    public IEnumerator<PdfExportWarning> GetEnumerator() => _warnings.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public override string ToString() => string.Join(Environment.NewLine, _warnings);

    /// <summary>
    /// Records a compromise. The same kind of compromise about the same subject is recorded
    /// once: a font missing from a thousand cells is one thing gone wrong, not a thousand.
    /// </summary>
    internal void Add(PdfExportWarningKind kind, string message, string? subject = null)
    {
        if (_seen.Add((kind, subject)))
            _warnings.Add(new PdfExportWarning(kind, message, subject));
    }
}
