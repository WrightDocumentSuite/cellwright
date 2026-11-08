namespace Cellwright.Pdf.Layout;

/// <summary>
/// A structure element a piece of the page belongs to, named by what it belongs to rather than
/// by position — so a header row asked for twice on two pages stays one element.
/// </summary>
/// <param name="Tag">The element type, as the logical structure names it.</param>
/// <param name="Owner">Whatever the element stands for, compared by reference.</param>
internal sealed record TagRef(string Tag, object Owner)
{
    /// <summary>The element this one hangs from, or <see langword="null"/> for a top-level element.</summary>
    public TagRef? Parent { get; init; }

    /// <summary>Text that replaces the content for a reader that cannot see it.</summary>
    public string? AlternateText { get; init; }

    /// <inheritdoc />
    public bool Equals(TagRef? other) =>
        other is not null && Tag == other.Tag && ReferenceEquals(Owner, other.Owner);

    /// <inheritdoc />
    public override int GetHashCode() =>
        HashCode.Combine(Tag, System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Owner));
}

/// <summary>
/// One page after composition: the sheet it came from, the paper it goes on, and everything to
/// be drawn on it.
/// </summary>
/// <remarks>
/// The furniture — headers and footers — is kept apart from the body because it is written first
/// and, when the page is scaled, it is not scaled with the body unless the sheet says so.
/// </remarks>
internal sealed class ComposedPage
{
    internal ComposedPage(Worksheet sheet, PageGeometry geometry, int index, int number)
    {
        Sheet = sheet;
        Geometry = geometry;
        Index = index;
        Number = number;
    }

    /// <summary>The sheet this page shows part of.</summary>
    public Worksheet Sheet { get; }

    /// <summary>The paper and its margins.</summary>
    public PageGeometry Geometry { get; }

    /// <summary>Where the page sits in the document, counting from zero.</summary>
    public int Index { get; }

    /// <summary>The number the page prints, which a sheet may restart.</summary>
    public int Number { get; set; }

    /// <summary>How many pages the sheet came to, which is what a page-count field prints.</summary>
    public int SheetPageCount { get; set; } = 1;

    /// <summary>The grid and everything in it, drawn under the print scale.</summary>
    public List<PageItem> Items { get; } = [];

    /// <summary>The headers and footers, drawn outside the print scale.</summary>
    public List<PageItem> Furniture { get; } = [];

    /// <summary>
    /// How much the body is scaled, as a fraction: a half means the grid is drawn at half size,
    /// so twice as much of it fits.
    /// </summary>
    public double Scale { get; set; } = 1;

    /// <summary>
    /// How far the scaled body is shifted right and down, which is what centring amounts to
    /// once the scale is known.
    /// </summary>
    public (double X, double Y) Offset { get; set; }

    /// <summary>The columns drawn on this page, or <see langword="null"/> for a blank page.</summary>
    public PageSlots? ColumnSlots { get; set; }

    /// <summary>The rows drawn on this page, or <see langword="null"/> for a blank page.</summary>
    public PageSlots? RowSlots { get; set; }
}
