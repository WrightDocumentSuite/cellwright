using Cellwright.Styles;

namespace Cellwright.Pdf.Layout;

/// <summary>
/// Headers and footers: which one a page gets, and where on the paper its three sections sit.
/// </summary>
/// <remarks>
/// They are laid out after the body because a header usually prints the page number, and the
/// page number is the one thing the body cannot know until the sheet has finished falling into
/// pages. They also sit outside the print scale, in the margin the page setup keeps for them.
/// </remarks>
internal sealed partial class SheetComposer
{
    /// <summary>The font Excel gives a header that names none.</summary>
    private static FontStyle HeaderFont => FontStyle.Default with { Size = 10 };

    /// <summary>Draws the header and the footer of every page, now that the numbers are known.</summary>
    private void Furnish(List<ComposedPage> pages)
    {
        if (_setup.Header is null && _setup.Footer is null)
            return;

        WarnAboutFurniture(pages.Count > 0 ? pages[0].Scale : 1);
        DateTime now = DateTime.Now;

        foreach (ComposedPage page in pages)
        {
            var fields = new HeaderFooterParser.Fields(page.Number, page.SheetPageCount, _sheet.Name, now);
            bool first = page.Index == _firstIndex;

            Draw(page, Which(_setup.Header, page, first), fields, footer: false);
            Draw(page, Which(_setup.Footer, page, first), fields, footer: true);
        }
    }

    /// <summary>
    /// Which of the three texts a page uses: its own for the first page or an even one, when the
    /// sheet turned those on, and the ordinary one otherwise.
    /// </summary>
    /// <remarks>
    /// A sheet that asks for a first page of its own and then leaves that text empty is asking
    /// for a first page with no header, which is how a title page is made. Falling back on the
    /// ordinary text would print the very thing the sheet asked to be rid of.
    /// </remarks>
    private string? Which(HeaderFooterText? text, ComposedPage page, bool first)
    {
        if (text is null)
            return null;

        if (first && _setup.DifferentFirstPage)
            return text.First;

        if (_setup.DifferentOddEven && page.Number % 2 == 0)
            return text.Even;

        return text.Default;
    }

    private void Draw(ComposedPage page, string? text, HeaderFooterParser.Fields fields, bool footer)
    {
        if (text is { Length: > 0 } && text.Contains("&G", StringComparison.Ordinal))
        {
            _context.Diagnostics.Add(
                PdfExportWarningKind.ContentSkipped,
                "A picture in a header or footer is not drawn.",
                "header-picture");
        }

        HeaderLine line = HeaderFooterParser.Parse(text, HeaderFont, fields);
        if (line.IsEmpty)
            return;

        foreach ((HeaderSection section, List<HeaderPiece> pieces) in line.Sections)
        {
            if (pieces.Count == 0)
                continue;

            Place(page, section, pieces, footer);
        }
    }

    /// <summary>Puts one section of one header or footer where the margins say.</summary>
    private void Place(ComposedPage page, HeaderSection section, List<HeaderPiece> pieces, bool footer)
    {
        var appearances = new List<TextAppearance>(pieces.Count);
        double width = 0;
        double tallest = 0;

        foreach (HeaderPiece piece in pieces)
        {
            TextAppearance appearance = _appearances.Of(piece.Font);
            appearances.Add(appearance);
            width += appearance.Measure(piece.Text);
            tallest = Math.Max(tallest, appearance.LineHeight);
        }

        double x = section switch
        {
            HeaderSection.Center => _geometry.ContentLeft + ((_geometry.ContentWidth - width) / 2),
            HeaderSection.Right => _geometry.ContentRight - width,
            _ => _geometry.ContentLeft,
        };

        // The header sits its own distance below the top of the paper; the footer that distance
        // above the bottom, with its own height taken off so it stays on the page.
        double y = footer
            ? _geometry.Height - _geometry.FooterDistance - tallest
            : _geometry.HeaderDistance;

        for (int i = 0; i < pieces.Count; i++)
        {
            double measured = appearances[i].Measure(pieces[i].Text);
            page.Furniture.Add(new TextItem
            {
                X = x,
                Y = y,
                Text = pieces[i].Text,
                Appearance = appearances[i],
                Width = measured,
                Height = appearances[i].LineHeight,
            });

            x += measured;
        }
    }
}
