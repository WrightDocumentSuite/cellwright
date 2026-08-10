namespace Cellwright.Pdf.Layout;

/// <summary>
/// Turns one sheet into pages: measures its grid, cuts the grid into pages that fit the paper,
/// and fills each page with what falls on it.
/// </summary>
internal sealed partial class SheetComposer
{
    private readonly PdfExportContext _context;
    private readonly Worksheet _sheet;
    private readonly PageSetup _setup;
    private readonly PageGeometry _geometry;
    private readonly int _firstIndex;

    private SheetTags? _tags;
    private SheetCommentPlan? _comments;
    private CellAppearances _appearances = null!;
    private ConditionalFormatEvaluator _conditional = null!;

    internal SheetComposer(PdfExportContext context, Worksheet sheet, int firstIndex)
    {
        _context = context;
        _sheet = sheet;
        _setup = sheet.PageSetup;
        _geometry = PageGeometry.From(sheet.PageSetup);
        _firstIndex = firstIndex;

        Warn();
    }

    /// <summary>Composes the pages the sheet comes to, or none when it has nothing to print.</summary>
    public List<ComposedPage> Compose()
    {
        SheetPlan? plan = SheetPlan.For(_context, _sheet);
        if (_context.Options.IncludeComments)
            _comments = new SheetCommentPlan(_context, _sheet);

        if (plan is null)
        {
            _comments?.Complete();
            return [];
        }

        if (_context.Options.Tagged)
            _tags = new SheetTags(_sheet, plan.TitleRows);

        var grid = new GridMeasurer(_context, _sheet, plan);
        _appearances = grid.Appearances;

        // One evaluator serves the whole sheet: its per-rule statistics are gathered once and
        // reused by every cell the rule covers, which a per-page evaluator would redo per page.
        _conditional = new ConditionalFormatEvaluator(
            _sheet,
            FormulaEvaluator(),
            evaluateRelative: RelativeFormulaEvaluator());

        List<ComposedPage> pages = Paginate(plan, grid);

        foreach (ComposedPage page in pages)
            page.SheetPageCount = pages.Count;

        Furnish(pages);
        ReportConditionalFormats(_conditional);
        _comments?.Complete();
        return pages;
    }

    /// <summary>The one empty page a workbook with nothing in it still prints.</summary>
    public List<ComposedPage> ComposeBlank() => [NewPage(0)];

    /// <summary>The caller's formula engine, bound to this sheet, or nothing when none was given.</summary>
    private Func<string, CellRef, CellValue>? FormulaEvaluator() =>
        _context.Options.FormulaEvaluator is { } evaluate
            ? (formula, cell) => evaluate(_sheet, cell, formula)
            : null;

    /// <summary>
    /// The caller's copy-aware formula engine, preserving the stored text as one cache key.
    /// </summary>
    private Func<string, CellRef, CellRef, CellValue>? RelativeFormulaEvaluator() =>
        _context.Options.RelativeFormulaEvaluator is { } evaluate
            ? (formula, anchor, cell) => evaluate(_sheet, anchor, cell, formula)
            : null;

    private ComposedPage NewPage(int within) => new(
        _sheet,
        _geometry,
        _firstIndex + within,
        (_setup.FirstPageNumber > 0 ? _setup.FirstPageNumber : 1) + within);
}
