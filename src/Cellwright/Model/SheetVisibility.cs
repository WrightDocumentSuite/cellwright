namespace Cellwright;

/// <summary>
/// How visible a sheet is in the Excel UI. The third state is not a stronger hidden: a sheet
/// hidden the ordinary way comes back through the tab context menu, while a very hidden one can
/// only be brought back from the VBA editor — which is why files use it to keep working sheets
/// out of a user's way for good.
/// </summary>
public enum SheetVisibility
{
    /// <summary>The sheet has a tab like any other.</summary>
    Visible,

    /// <summary>Hidden, and shown again from the tab context menu.</summary>
    Hidden,

    /// <summary>Hidden, and only shown again from the VBA editor.</summary>
    VeryHidden,
}
