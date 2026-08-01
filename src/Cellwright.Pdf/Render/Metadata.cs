using Inkwright;
using Inkwright.Cos;
using Inkwright.Metadata;

namespace Cellwright.Pdf.Render;

/// <summary>What the document says about itself: the properties Excel shows, in PDF's own terms.</summary>
internal static class Metadata
{
    /// <summary>Copies the workbook's properties onto the document.</summary>
    /// <param name="context">The render in progress.</param>
    public static void Apply(PdfExportContext context)
    {
        WorkbookProperties properties = context.Source.Properties;

        if (!string.IsNullOrWhiteSpace(properties.Title))
        {
            context.Pdf.Info.Title = properties.Title;

            // A tagged document is meant to be read out as well as looked at, and a reader that
            // announces the file name instead of the title is the first thing to go wrong.
            if (context.Options.Tagged)
                DisplayTitle(context.Pdf);
        }

        if (!string.IsNullOrWhiteSpace(properties.Author))
            context.Pdf.Info.Author = properties.Author;

        if (!string.IsNullOrWhiteSpace(properties.Subject))
            context.Pdf.Info.Subject = properties.Subject;

        if (!string.IsNullOrWhiteSpace(properties.Keywords))
            context.Pdf.Info.Keywords = properties.Keywords;

        if (properties.Created is { } created)
            context.Pdf.Info.CreationDate = created;

        WriteXmp(context);
    }

    /// <summary>
    /// Writes the metadata a second time, as XMP. The information dictionary is what old readers
    /// look at and XMP is what everything since PDF 1.4 looks at, including every archival and
    /// accessibility profile, so a document that carries only the first is half described.
    /// </summary>
    private static void WriteXmp(PdfExportContext context)
    {
        var metadata = new XmpMetadata();
        metadata.CopyFrom(context.Pdf.Info);

        if (context.Options.Language is { Length: > 0 } language)
            metadata.Language = language;

        metadata.ApplyTo(context.Pdf);
    }

    /// <summary>Asks readers to show the document's title rather than its file name.</summary>
    private static void DisplayTitle(PdfDocument pdf)
    {
        PdfName key = PdfName.Get("ViewerPreferences");
        PdfDictionary preferences = pdf.Catalog.GetDictionary(key) ?? new PdfDictionary(1);

        preferences.Set(PdfName.Get("DisplayDocTitle"), PdfValue.True);
        pdf.Catalog.Set(key, PdfValue.Dictionary(preferences));
        pdf.MarkChanged(pdf.CatalogId);
    }
}
