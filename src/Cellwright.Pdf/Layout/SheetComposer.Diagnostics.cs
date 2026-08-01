namespace Cellwright.Pdf.Layout;

/// <summary>
/// What a sheet carries that its printing does not.
/// </summary>
/// <remarks>
/// Everything left out is said once, when the sheet is taken up, rather than at the moment it
/// would have been drawn: a rule that is never reached and a comment that is never painted leave
/// no trace of themselves in the drawing, and a caller comparing the PDF with the workbook needs
/// to be told about both. The diagnostics keep one of each, so this costs one entry a sheet
/// however much the sheet holds.
/// </remarks>
internal sealed partial class SheetComposer
{
    private void Warn()
    {
        if (!SheetUnits.TryPaper(_setup.PaperSize, out _))
        {
            Say(
                PdfExportWarningKind.LayoutApproximated,
                "A paper size the format does not name was printed on A4.",
                $"paper-{_setup.PaperSize}");
        }

        if (_sheet.Comments.Count > 0)
            Say(PdfExportWarningKind.ContentSkipped, "A comment on a cell is not printed.", "comment");

        // Conditional formats are worked out rather than refused wholesale. Only the parts that
        // could not be applied are reported, and those are reported where they are met — see
        // ReportConditionalFormats, which runs after the cells have been judged.

        if (_setup.PrintHeadings)
        {
            Say(
                PdfExportWarningKind.ContentSkipped,
                "The row numbers and column letters a sheet asks to have printed are not drawn.",
                "print-headings");
        }

        if (_setup.Draft)
            Say(PdfExportWarningKind.LayoutApproximated, "A sheet set to print in draft is printed in full.", "draft");

        if (_setup.BlackAndWhite)
        {
            Say(
                PdfExportWarningKind.LayoutApproximated,
                "A sheet set to print without colour is printed in colour.",
                "black-and-white");
        }
    }

    /// <summary>What the headers and footers of this sheet cannot honour.</summary>
    /// <param name="scale">The scale the body of the sheet is drawn at.</param>
    private void WarnAboutFurniture(double scale)
    {
        if (_setup.ScaleHeaderWithDocument && scale is > 0 and not 1)
        {
            Say(
                PdfExportWarningKind.LayoutApproximated,
                "A header or footer is drawn at full size on a sheet whose body is scaled.",
                "header-scale");
        }

        if (!_setup.AlignHeaderWithMargins)
        {
            Say(
                PdfExportWarningKind.LayoutApproximated,
                "A header or footer is lined up with the page margins whatever the sheet asks.",
                "header-margins");
        }
    }

    private void Say(PdfExportWarningKind kind, string message, string subject) =>
        _context.Diagnostics.Add(kind, message, subject);

    /// <summary>
    /// Reports the parts of the sheet's conditional formatting that could not be applied, once
    /// the cells have been judged and it is known which parts those were.
    /// </summary>
    /// <remarks>
    /// Said afterwards rather than up front on purpose: which rules defeat the renderer depends
    /// on what is in them, and a sheet whose rules were all applied should say nothing at all.
    /// </remarks>
    private void ReportConditionalFormats(ConditionalFormatEvaluator evaluator)
    {
        foreach (string subject in evaluator.Unsupported)
        {
            Say(
                PdfExportWarningKind.LayoutApproximated,
                subject switch
                {
                    "conditional-format-formula" =>
                        "A conditional format whose condition is a formula was not worked out: no formula " +
                        "evaluator was supplied. Set PdfExportOptions.FormulaEvaluator to have it applied.",
                    "conditional-format-opaque" =>
                        "A conditional formatting rule the model does not represent was not worked out; " +
                        "the cell is printed in the style it states.",
                    _ => "Part of a conditional formatting rule was not worked out; the cell is printed " +
                        "in the style it states.",
                },
                subject);
        }
    }
}
