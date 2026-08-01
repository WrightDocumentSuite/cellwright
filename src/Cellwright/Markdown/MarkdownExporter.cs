using System.Text;

namespace Cellwright.Markdown;

/// <summary>
/// Turns a workbook into one Markdown document plus the image files it refers to.
/// </summary>
/// <remarks>
/// A pure model-to-text conversion: nothing here opens a file or a package, which is why it sits
/// in the core library beside the writers rather than in a format-specific one. The whole
/// document is built in memory — the images have to be handed over as a sidecar anyway, so
/// streaming the text would buy nothing.
/// </remarks>
internal static class MarkdownExporter
{
    /// <summary>Heading of the picture gallery under a sheet's table.</summary>
    private const string PicturesHeading = "Изображения";

    /// <summary>What a sheet with nothing on it reads as, so that its heading is not left bare.</summary>
    private const string EmptySheet = "(пусто)";

    /// <summary>Title of a workbook whose properties name none.</summary>
    private const string UntitledWorkbook = "Workbook";

    /// <summary>Converts the workbook. Hidden sheets are left out, as they are out of sight in Excel.</summary>
    public static MarkdownDocument Export(Workbook workbook, MarkdownExportOptions options)
    {
        string mediaDirectoryName = MarkdownPath.NormalizeMediaDirectory(options.MediaDirectoryName);
        options = options with { MediaDirectoryName = mediaDirectoryName };

        var builder = new StringBuilder();
        var media = new MarkdownMedia();

        builder.Append("# ");
        MarkdownText.AppendHeading(builder, Title(workbook, options));
        builder.Append("\n\n");

        foreach (Worksheet sheet in workbook.Sheets)
        {
            if (sheet.Visibility == SheetVisibility.Visible)
                WriteSheet(builder, sheet, options, media);
        }

        return new MarkdownDocument(Finish(builder), media.Images, options.MediaDirectoryName);
    }

    private static void WriteSheet(
        StringBuilder builder, Worksheet sheet, MarkdownExportOptions options, MarkdownMedia media)
    {
        builder.Append("## ");
        MarkdownText.AppendHeading(builder, sheet.Name);
        builder.Append("\n\n");

        if (sheet.UsedRange is { } used)
            MarkdownTable.Write(builder, sheet, used, options);
        else
            builder.Append(EmptySheet).Append("\n\n");

        if (options.IncludePictures)
            WriteGallery(builder, sheet, options, media);
    }

    /// <summary>
    /// Lists a sheet's pictures under its table. A GFM table holds neither a spanning cell nor a
    /// floating object, so an image cannot go where it sits on the sheet; a gallery in reading
    /// order of the anchors is the closest a document gets.
    /// </summary>
    private static void WriteGallery(
        StringBuilder builder, Worksheet sheet, MarkdownExportOptions options, MarkdownMedia media)
    {
        List<Picture> shown =
        [
            .. sheet.Pictures
                .Where(static picture => !picture.IsHidden)
                .OrderBy(static picture => picture.Anchor.From.Row)
                .ThenBy(static picture => picture.Anchor.From.Column),
        ];

        if (shown.Count == 0)
            return;

        builder.Append("### ").Append(PicturesHeading).Append("\n\n");
        foreach (Picture picture in shown)
        {
            string file = media.Add(picture);
            builder.Append('!');
            MarkdownText.AppendLink(
                builder, AltText(picture, file), MarkdownPath.MediaUrl(options.MediaDirectoryName, file));
            builder.Append("\n\n");
        }
    }

    /// <summary>
    /// Alternative text of a picture. The description comes first: it is the alternative text
    /// the file was given, while the name is usually whatever Excel numbered the shape.
    /// </summary>
    private static string AltText(Picture picture, string fileName)
    {
        if (picture.Description is { Length: > 0 } description)
            return description;
        return picture.Name is { Length: > 0 } name ? name : fileName;
    }

    private static string Title(Workbook workbook, MarkdownExportOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.WorkbookTitle))
            return options.WorkbookTitle;
        return string.IsNullOrWhiteSpace(workbook.Properties.Title) ? UntitledWorkbook : workbook.Properties.Title;
    }

    /// <summary>Closes the document with exactly one newline, however many the last block left.</summary>
    private static string Finish(StringBuilder builder)
    {
        while (builder.Length > 0 && builder[builder.Length - 1] == '\n')
            builder.Length--;

        builder.Append('\n');
        return builder.ToString();
    }
}
