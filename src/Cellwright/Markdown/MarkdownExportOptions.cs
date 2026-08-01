using System.Globalization;

namespace Cellwright;

/// <summary>
/// What the header row of a sheet's table is made of. GitHub Flavored Markdown has no
/// headerless table, so one has to be chosen.
/// </summary>
public enum MarkdownHeaderRow
{
    /// <summary>The first row of the used range becomes the header and leaves the data.</summary>
    FirstRow,

    /// <summary>Column letters (<c>A</c>, <c>B</c>, …) become the header and every row stays data.</summary>
    ColumnLetters,
}

/// <summary>
/// Options for <see cref="Workbook.ToMarkdown(MarkdownExportOptions?)"/> and
/// <see cref="Workbook.ExportMarkdownAsync(string, MarkdownExportOptions?, CancellationToken)"/>.
/// </summary>
public sealed record MarkdownExportOptions
{
    internal static readonly MarkdownExportOptions Default = new();

    /// <summary>
    /// Culture for separators and month names while cells are rendered through their number
    /// formats. Invariant by default, like the rest of the library's rendering: an export is a
    /// document, and a document that reads differently on another machine is a bug.
    /// </summary>
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>Where the header row of each table comes from.</summary>
    public MarkdownHeaderRow HeaderRow { get; init; } = MarkdownHeaderRow.FirstRow;

    /// <summary>Whether a sheet's pictures are collected and listed under its table.</summary>
    public bool IncludePictures { get; init; } = true;

    /// <summary>
    /// Directory the image links point into, relative to the document. Directory separators are
    /// normalized; rooted paths, parent traversal and values without a directory name are rejected.
    /// </summary>
    public string MediaDirectoryName { get; init; } = "media";

    /// <summary>
    /// Title of the top-level heading. Falls back to <see cref="WorkbookProperties.Title"/>
    /// and then to <c>Workbook</c>.
    /// </summary>
    public string? WorkbookTitle { get; init; }
}
