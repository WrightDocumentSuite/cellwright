namespace Cellwright;

/// <summary>
/// The locked and hidden flags of a cell style. Neither does anything until the sheet itself is
/// protected — which is why the default is locked: protecting a sheet is meant to freeze it, and
/// the cells people may still fill in are the ones explicitly unlocked beforehand.
/// </summary>
/// <example>
/// <code>
/// // An input cell on a sheet that will be protected
/// sheet.Cell("B2").SetStyle(CellStyle.Default with { Protection = CellProtection.Unlocked });
/// </code>
/// </example>
public sealed record CellProtection
{
    /// <summary>Excel's default: locked, formula visible.</summary>
    public static CellProtection Default { get; } = new();

    /// <summary>A cell that stays editable on a protected sheet.</summary>
    public static CellProtection Unlocked { get; } = new() { Locked = false };

    /// <summary>Whether protecting the sheet freezes this cell.</summary>
    public bool Locked { get; init; } = true;

    /// <summary>Whether protecting the sheet hides this cell's formula from the formula bar.</summary>
    public bool Hidden { get; init; }
}
