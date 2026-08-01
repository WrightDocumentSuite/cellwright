using System.Globalization;

namespace Cellwright.Formulas;

/// <summary>
/// Options for a calculation pass.
/// </summary>
public sealed record CalculateOptions
{
    internal static readonly CalculateOptions Default = new();

    /// <summary>
    /// Recalculates every formula instead of only the ones affected by changes since the last
    /// pass. Volatile formulas are recalculated either way.
    /// </summary>
    public bool Force { get; init; }

    /// <summary>
    /// Culture whose separators the format string of <c>TEXT</c> is written in — the locale of
    /// the machine that authored the file. Excel stores those strings as the user typed them,
    /// so a Russian workbook says <c>TEXT(A1;"# ##0,00")</c> where an invariant reading of the
    /// comma and the space gets both wrong. The output uses the same culture's separators.
    /// Defaults to invariant, which reads formats the way this documentation writes them.
    /// Localized date letters (<c>ДД.ММ</c> for <c>dd.mm</c>) are not translated.
    /// </summary>
    public CultureInfo? FormatCulture { get; init; }

    /// <summary>
    /// Called for each recoverable problem: circular references, unparsable formulas,
    /// references to missing sheets.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; init; }

    /// <summary>Cancels a long recalculation; checked between cells.</summary>
    public CancellationToken CancellationToken { get; init; }
}

/// <summary>
/// What a calculation pass did.
/// </summary>
/// <param name="CalculatedCells">Number of formula cells evaluated.</param>
/// <param name="ErrorCells">How many of them ended in an Excel error value.</param>
/// <param name="CircularCells">Cells that take part in a circular reference; each holds zero.</param>
/// <param name="Elapsed">Wall-clock duration of the pass.</param>
public sealed record CalculationSummary(
    int CalculatedCells,
    int ErrorCells,
    IReadOnlyList<string> CircularCells,
    TimeSpan Elapsed)
{
    /// <summary>Whether any circular reference was found.</summary>
    public bool HasCircularReferences => CircularCells.Count > 0;

    /// <inheritdoc/>
    public override string ToString() =>
        $"{CalculatedCells} cell(s) in {Elapsed.TotalMilliseconds:F1} ms, {ErrorCells} error(s), {CircularCells.Count} circular";
}
