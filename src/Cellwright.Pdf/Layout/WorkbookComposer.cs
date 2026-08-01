namespace Cellwright.Pdf.Layout;

/// <summary>
/// Turns a workbook into pages: which sheets are printed, in what order, and what each of them
/// comes to.
/// </summary>
/// <remarks>
/// Page numbers run through the whole document unless a sheet restarts them, which is what makes
/// a header printing "page 3 of 12" mean the same thing here as in Excel.
/// </remarks>
internal sealed class WorkbookComposer
{
    private readonly PdfExportContext _context;

    internal WorkbookComposer(PdfExportContext context) => _context = context;

    /// <summary>Composes every sheet the options select.</summary>
    public IReadOnlyList<ComposedPage> Compose()
    {
        List<Worksheet> selected = [.. Selected()];
        List<ComposedPage> pages = [];

        foreach (Worksheet sheet in selected)
        {
            var composer = new SheetComposer(_context, sheet, pages.Count);
            List<ComposedPage> sheetPages = composer.Compose();

            // The sheet has checked its own budget already; this catches what a blank page or a
            // future kind of sheet might add on top of it.
            PageBudget.Check(pages.Count + sheetPages.Count, _context.Options);
            pages.AddRange(sheetPages);
        }

        // A workbook whose sheets are all empty still prints one page, the way Excel prints one —
        // and a PDF with no pages at all is not a document a reader will open.
        if (pages.Count == 0)
            pages.AddRange(new SheetComposer(_context, Blank(selected), 0).ComposeBlank());

        return pages;
    }

    /// <summary>The sheet a workbook with nothing to print takes its one page's paper from.</summary>
    /// <param name="selected">The sheets that were asked for, whether or not they printed.</param>
    /// <remarks>
    /// The sheet that was asked for is the one whose paper the caller meant, even though it has
    /// nothing on it: printing an empty landscape sheet on portrait paper because it happens to
    /// be the second in the workbook is not what was asked for. A workbook with no sheets at all
    /// borrows one from a workbook of our own — rendering is a read of the caller's document and
    /// must leave it as it found it.
    /// </remarks>
    private Worksheet Blank(List<Worksheet> selected)
    {
        if (selected.Count > 0)
            return selected[0];

        if (_context.Source.Sheets.Count > 0)
            return _context.Source.Sheets[0];

        Workbook scratch = Workbook.Create();
        return scratch.Sheets.Count > 0 ? scratch.Sheets[0] : scratch.Sheets.Add("Sheet1");
    }

    /// <summary>The sheets to print, in the order they are printed.</summary>
    private IEnumerable<Worksheet> Selected()
    {
        if (_context.Options.Selection == SheetSelection.Named)
        {
            foreach (string name in _context.Options.Sheets)
            {
                Worksheet? sheet = _context.Source.Sheets.FirstOrDefault(
                    candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));

                if (sheet is null)
                {
                    _context.Diagnostics.Add(
                        PdfExportWarningKind.ContentSkipped,
                        "A sheet named for printing is not in the workbook.",
                        name);
                    continue;
                }

                yield return sheet;
            }

            yield break;
        }

        bool all = _context.Options.Selection == SheetSelection.All;
        foreach (Worksheet sheet in _context.Source.Sheets)
        {
            if (all || !sheet.IsHidden)
                yield return sheet;
        }
    }
}
