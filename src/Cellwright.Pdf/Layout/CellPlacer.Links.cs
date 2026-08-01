namespace Cellwright.Pdf.Layout;

/// <summary>The clickable areas a sheet's hyperlinks put over its cells.</summary>
internal sealed partial class CellPlacer
{
    /// <summary>Puts a clickable area over every cell a hyperlink covers on this page.</summary>
    public void Links()
    {
        foreach (Hyperlink link in _sheet.Hyperlinks)
        {
            if (link.Target is not { Length: > 0 } url)
            {
                // A link to a place inside the workbook has nowhere to point once the workbook
                // is a stack of pages: the sheet it names may not even be printed.
                _context.Diagnostics.Add(
                    PdfExportWarningKind.ContentSkipped,
                    "A hyperlink that points inside the workbook is not clickable in the PDF.",
                    "internal-link");
                continue;
            }

            Cover(link.Range, url);
        }
    }

    /// <summary>Covers whatever part of a link's range this page shows.</summary>
    private void Cover(RangeRef range, string url)
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            int row = _rows.Lines[r];
            if (row < range.Start.Row || row > range.End.Row)
                continue;

            for (int c = 0; c < _columns.Count; c++)
            {
                int column = _columns.Lines[c];
                if (column < range.Start.Column || column > range.End.Column)
                    continue;

                _page.Items.Add(new LinkItem
                {
                    X = _columns.Near[c],
                    Y = _rows.Near[r],
                    Width = _columns.Size[c],
                    Height = _rows.Size[r],
                    Url = url,
                });
            }
        }
    }
}
