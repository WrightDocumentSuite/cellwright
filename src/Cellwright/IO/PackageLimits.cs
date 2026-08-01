namespace Cellwright;

/// <summary>
/// What a package is allowed to cost while it is being read. A zip declares how far each of its
/// entries expands and is under no obligation to be honest about it, so a small file built for
/// the purpose can ask for as much memory as it likes. These bounds turn that into a message
/// instead of an <see cref="OutOfMemoryException"/>.
/// </summary>
/// <remarks>
/// Every part is checked against <see cref="MaxPartSize"/> as it is decompressed. Only parts held
/// in memory are added to <see cref="MaxTotalSize"/>: styles, shared strings, notes, and everything
/// carried through verbatim. Worksheets are streamed through the reader and never buffered,
/// however large they are, so a workbook of millions of rows is unaffected by the total budget.
/// </remarks>
public sealed record PackageLimits
{
    /// <summary>Bounds generous enough for any real workbook and far short of exhausting a server.</summary>
    public static PackageLimits Default { get; } = new();

    /// <summary>No bounds at all, for a caller who knows where the file came from.</summary>
    public static PackageLimits None { get; } = new() { MaxPartSize = long.MaxValue, MaxTotalSize = long.MaxValue };

    /// <summary>Largest decompressed single part accepted by the reader. Defaults to 512 MiB.</summary>
    public long MaxPartSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Largest total the buffered parts may reach together. Defaults to 2 GiB.</summary>
    public long MaxTotalSize { get; init; } = 2L * 1024 * 1024 * 1024;
}
