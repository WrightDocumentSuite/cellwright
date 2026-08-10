namespace Cellwright;

/// <summary>
/// Compatibility view of the original package-only limits. New code should use
/// <see cref="DocumentLoadBudget"/>, which also bounds entry counts, XML, cells and resources.
/// </summary>
/// <remarks>
/// Setting this facade on a load option maps <see cref="MaxPartSize"/> to the per-part ceiling and
/// <see cref="MaxTotalSize"/> to both source bytes and total inflated payload. The rest of the
/// finite default document budget stays in force, except that <see cref="None"/> deliberately
/// restores the old fully-unlimited behavior. When an asymmetric <see cref="DocumentLoadBudget"/>
/// is viewed through this facade, <see cref="MaxTotalSize"/> reports the tighter of its source and
/// inflated ceilings so a compatibility round trip cannot loosen either limit.
/// </remarks>
public sealed record PackageLimits
{
    /// <summary>Bounds generous enough for any real workbook and far short of exhausting a server.</summary>
    public static PackageLimits Default { get; } = new();

    /// <summary>No bounds at all, for a caller who knows where the file came from.</summary>
    public static PackageLimits None { get; } = new() { MaxPartSize = long.MaxValue, MaxTotalSize = long.MaxValue };

    /// <summary>Largest decompressed single part accepted by the reader. Defaults to 512 MiB.</summary>
    public long MaxPartSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Largest source and total uncompressed payload accepted. Defaults to 2 GiB.</summary>
    public long MaxTotalSize { get; init; } = 2L * 1024 * 1024 * 1024;

    /// <summary>Projects these legacy two-value limits onto the shared load budget.</summary>
    internal DocumentLoadBudget ApplyTo(DocumentLoadBudget budget) =>
        MaxPartSize == long.MaxValue && MaxTotalSize == long.MaxValue
            ? DocumentLoadBudget.Unlimited
            : budget with
            {
                MaxInputBytes = MaxTotalSize,
                MaxInflatedBytes = MaxTotalSize,
                MaxPartBytes = MaxPartSize,
            };

    /// <summary>Creates the compatibility view exposed by older options APIs.</summary>
    internal static PackageLimits From(DocumentLoadBudget budget) => new()
    {
        MaxPartSize = budget.MaxPartBytes,
        // One legacy value represented both ceilings. Projecting the tighter one keeps a
        // Budget -> Limits -> Budget round trip from silently weakening either bound.
        MaxTotalSize = Math.Min(budget.MaxInputBytes, budget.MaxInflatedBytes),
    };
}
