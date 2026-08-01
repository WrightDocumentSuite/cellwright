using Inkwright;
using Inkwright.Text;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Renders a workbook and reads the result back the way a viewer would: through a saved file, so
/// that anything the writer defers — font subsetting, the <c>ToUnicode</c> map — has happened.
/// </summary>
internal sealed class Rendered : IDisposable
{
    private Rendered(PdfDocument document, PdfExportDiagnostics diagnostics)
    {
        Document = document;
        Diagnostics = diagnostics;
    }

    public PdfDocument Document { get; }

    public PdfExportDiagnostics Diagnostics { get; }

    public int PageCount => Document.Pages.Count;

    /// <summary>Renders a workbook and reads the result back.</summary>
    /// <param name="workbook">The workbook to render.</param>
    /// <param name="options">How to render it.</param>
    /// <param name="before">
    /// What to do to the rendered document before it is saved, which is where a caller would
    /// sign it, encrypt it or claim conformance with a profile.
    /// </param>
    public static Rendered Of(
        Workbook workbook, PdfExportOptions? options = null, Action<PdfDocument>? before = null)
    {
        PdfExportResult result = PdfExporter.Render(workbook, options);
        byte[] bytes;
        using (result.Document)
        {
            before?.Invoke(result.Document);
            bytes = result.Document.ToArray();
        }

        return new Rendered(PdfDocument.Load(bytes), result.Diagnostics);
    }

    public string Text(int page = 0) => Document.Pages[page].ExtractText();

    public IReadOnlyList<PdfLetter> Letters(int page = 0) => Document.Pages[page].ExtractLetters();

    /// <summary>The content stream of a page, as the operators a test can look for.</summary>
    public string Content(int page = 0) =>
        System.Text.Encoding.Latin1.GetString(Document.Pages[page].GetContent());

    /// <summary>Where a piece of text was drawn, or null when the page does not show it.</summary>
    /// <param name="text">The text to find, as one run of letters.</param>
    /// <param name="page">The page to look at.</param>
    public (double X, double Y)? Find(string text, int page = 0)
    {
        IReadOnlyList<PdfLetter> letters = Letters(page);

        for (int i = 0; i + text.Length <= letters.Count; i++)
        {
            bool match = true;
            for (int j = 0; j < text.Length && match; j++)
                match = letters[i + j].Text == text[j].ToString();

            if (match)
                return (letters[i].Origin.X, letters[i].Origin.Y);
        }

        return null;
    }

    /// <summary>
    /// The lines of a page as a reader sees them, top down. Glyphs within a few points of each
    /// other vertically belong to one line, which is what keeps a row of a table together.
    /// </summary>
    public IReadOnlyList<string> Lines(int page = 0)
    {
        const double SameLine = 4;

        List<string> lines = [];
        List<PdfLetter> current = [];
        double previous = double.NaN;

        foreach (PdfLetter letter in Letters(page).OrderByDescending(letter => letter.Origin.Y))
        {
            if (current.Count > 0 && previous - letter.Origin.Y > SameLine)
            {
                lines.Add(Join(current));
                current.Clear();
            }

            current.Add(letter);
            previous = letter.Origin.Y;
        }

        if (current.Count > 0)
            lines.Add(Join(current));

        return lines;

        static string Join(List<PdfLetter> letters) =>
            string.Concat(letters.OrderBy(letter => letter.Origin.X).Select(letter => letter.Text));
    }

    public void Dispose() => Document.Dispose();
}
