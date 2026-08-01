namespace Cellwright;

/// <summary>
/// One image a Markdown export refers to: the bytes of the picture and the file name its links
/// were written with. The bytes are the picture's own, handed over without re-encoding.
/// </summary>
public sealed class MarkdownImage
{
    /// <summary>File name inside the media directory, e.g. <c>image1.png</c>.</summary>
    public required string FileName { get; init; }

    /// <summary>Encoded image, exactly as the workbook stored it.</summary>
    public required ReadOnlyMemory<byte> Content { get; init; }

    /// <summary>Encoding of <see cref="Content"/>.</summary>
    public required PictureFormat Format { get; init; }

    /// <inheritdoc/>
    public override string ToString() => $"{FileName} ({Content.Length} bytes)";
}

/// <summary>
/// The result of turning a workbook into Markdown: the text of the document and the images its
/// links point at. Nothing is written anywhere until <see cref="SaveAsync"/> is called, so the
/// caller is free to put the two halves wherever they belong — a static site, a wiki, a request.
/// </summary>
public sealed class MarkdownDocument
{
    /// <summary>Name <see cref="SaveAsync"/> gives the document.</summary>
    public const string DefaultFileName = "workbook.md";

    internal MarkdownDocument(string text, IReadOnlyList<MarkdownImage> images, string mediaDirectoryName)
    {
        Text = text;
        Images = images;
        MediaDirectoryName = Markdown.MarkdownPath.NormalizeMediaDirectory(mediaDirectoryName);
    }

    /// <summary>The Markdown text, with <c>\n</c> line endings.</summary>
    public string Text { get; }

    /// <summary>Images the text links to, in the order they were first referenced.</summary>
    public IReadOnlyList<MarkdownImage> Images { get; }

    /// <summary>Normalized relative directory the image links are relative to.</summary>
    public string MediaDirectoryName { get; }

    /// <summary>
    /// Writes <see cref="DefaultFileName"/> and the media directory into
    /// <paramref name="directoryPath"/>, creating both if they are not there. Existing files of
    /// the same names are overwritten; nothing else in the directory is touched.
    /// </summary>
    public async ValueTask SaveAsync(string directoryPath, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(directoryPath);
        Directory.CreateDirectory(directoryPath);
        await File.WriteAllTextAsync(Path.Combine(directoryPath, DefaultFileName), Text, cancellationToken)
            .ConfigureAwait(false);

        if (Images.Count == 0)
            return;

        string media = Path.Combine(directoryPath, MediaDirectoryName);
        Directory.CreateDirectory(media);
        foreach (MarkdownImage image in Images)
        {
            await File.WriteAllBytesAsync(Path.Combine(media, image.FileName), image.Content, cancellationToken)
                .ConfigureAwait(false);
        }
    }

    /// <inheritdoc/>
    public override string ToString() => Text;
}
