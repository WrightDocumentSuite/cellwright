using System.Security.Cryptography;
using Cellwright.Streaming;

namespace Cellwright.Markdown;

/// <summary>
/// The image files a Markdown export needs beside its text, collected as the sheets are walked.
/// </summary>
/// <remarks>
/// Files are keyed by the digest of their content rather than by the picture that asked for
/// them: the same logo anchored on three sheets is one file and three links. Bytes are passed
/// through untouched, as they are everywhere else in the library.
/// </remarks>
internal sealed class MarkdownMedia
{
    private readonly Dictionary<string, string> _namesByDigest = new(StringComparer.Ordinal);
    private readonly List<MarkdownImage> _images = [];

    /// <summary>Files to write, in the order the document first referred to them.</summary>
    public IReadOnlyList<MarkdownImage> Images => _images;

    /// <summary>Returns the file name a picture's content is published under, adding it if new.</summary>
    public string Add(Picture picture)
    {
        string digest = Convert.ToHexStringLower(SHA256.HashData(picture.Content.Span));
        if (_namesByDigest.TryGetValue(digest, out string? existing))
            return existing;

        string name = $"image{_images.Count + 1}.{DrawingWriter.Extension(picture.Format)}";
        _namesByDigest.Add(digest, name);
        _images.Add(new MarkdownImage
        {
            FileName = name,
            Content = picture.Content,
            Format = picture.Format,
        });
        return name;
    }
}
