using System.Globalization;

namespace Cellwright.Pdf;

/// <summary>Which sheets of a workbook are printed.</summary>
public enum SheetSelection
{
    /// <summary>Every sheet the workbook shows, in tab order.</summary>
    Visible,

    /// <summary>Every sheet, hidden ones included.</summary>
    All,

    /// <summary>Only the sheets named in <see cref="PdfExportOptions.Sheets"/>.</summary>
    Named,
}

/// <summary>
/// How a workbook is rendered. Everything has a working default: the options are for the
/// decisions a caller has to make differently, not for the ones it always makes the same way.
/// </summary>
public sealed class PdfExportOptions
{
    /// <summary>Which sheets are printed.</summary>
    public SheetSelection Selection { get; set; } = SheetSelection.Visible;

    /// <summary>
    /// The sheets to print when <see cref="Selection"/> is <see cref="SheetSelection.Named"/>,
    /// in the order they are named.
    /// </summary>
    public IList<string> Sheets { get; } = [];

    /// <summary>
    /// Faces to draw a family with, by family name — optionally with a style, as in
    /// <c>"Calibri Bold"</c>. A machine with no fonts installed needs these; anything not named
    /// here is looked for among the installed fonts.
    /// </summary>
    public IDictionary<string, string> FontFiles { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>The family drawn when the one a cell asks for is nowhere to be found.</summary>
    public string FallbackFontFamily { get; set; } = "Arial";

    /// <summary>Whether embedded faces carry only the glyphs the workbook used.</summary>
    public bool SubsetFonts { get; set; } = true;

    /// <summary>
    /// The culture number formats are rendered in — the separators, the month names. The
    /// invariant culture by default, so a document renders the same on every machine.
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// How to work out a conditional formatting rule whose condition is a formula, or
    /// <see langword="null"/> to leave such rules unapplied and report them.
    /// </summary>
    /// <remarks>
    /// Nearly every kind of rule is decided from the values alone and needs nothing here: the
    /// comparisons against literals, the text tests, the ranks, the averages, the colour scales,
    /// the bars and the icons are all applied whether or not this is set. Only a rule holding a
    /// real formula needs an engine, and this package deliberately does not depend on one — a
    /// caller that wants those rules applied says so by handing the engine over:
    /// <code>
    /// options.FormulaEvaluator = (sheet, cell, formula) =&gt; sheet.Evaluate(formula, cell.ToString());
    /// </code>
    /// which needs the <c>Cellwright.Formulas</c> package. Left unset, such a rule is reported as
    /// <c>conditional-format-formula</c> rather than quietly treated as false.
    /// </remarks>
    public Func<Worksheet, CellRef, string, CellValue>? FormulaEvaluator { get; set; }

    /// <summary>Whether the structure tree is built, which is what a screen reader follows.</summary>
    public bool Tagged { get; set; }

    /// <summary>
    /// The language of the document, as a BCP 47 tag. Only meaningful when
    /// <see cref="Tagged"/> is set, and required for accessibility conformance.
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    /// The most pages one render may produce. A sheet that cannot fit its content would
    /// otherwise paginate for ever; this turns that into an exception a caller can act on.
    /// </summary>
    public int MaxPages { get; set; } = 10_000;
}
