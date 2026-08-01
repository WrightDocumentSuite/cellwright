namespace Cellwright.Xls;

/// <summary>
/// Settings for writing a legacy workbook.
/// </summary>
public sealed class XlsSaveOptions
{
    /// <summary>
    /// Called for every thing the old format cannot carry — a sheet past its 65536×256 grid, a
    /// formula with no token for it, a note. Nothing here stops the save; the file is written
    /// with the closest thing the format has.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; set; }

    /// <summary>Whether to leave the destination stream open after writing.</summary>
    public bool LeaveOpen { get; set; }
}
