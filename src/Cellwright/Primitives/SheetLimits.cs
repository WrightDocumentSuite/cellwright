namespace Cellwright;

/// <summary>
/// Hard limits of the xlsx worksheet grid and related workbook constraints.
/// </summary>
public static class SheetLimits
{
    /// <summary>Maximum 1-based row number of a worksheet (2^20).</summary>
    public const int MaxRow = 1_048_576;

    /// <summary>Maximum 1-based column number of a worksheet (2^14, column "XFD").</summary>
    public const int MaxColumn = 16_384;

    /// <summary>Maximum length of a worksheet name.</summary>
    public const int MaxSheetNameLength = 31;

    /// <summary>Maximum number of characters a single cell can hold.</summary>
    public const int MaxCellTextLength = 32_767;

    /// <summary>Maximum column width in characters. Excel rejects a file that exceeds it.</summary>
    public const double MaxColumnWidth = 255.0;

    /// <summary>Maximum row height in points.</summary>
    public const double MaxRowHeight = 409.5;

    /// <summary>Returns <see langword="true"/> when the 1-based row number lies inside the grid.</summary>
    public static bool IsValidRow(int row) => row is >= 1 and <= MaxRow;

    /// <summary>Returns <see langword="true"/> when the 1-based column number lies inside the grid.</summary>
    public static bool IsValidColumn(int column) => column is >= 1 and <= MaxColumn;
}
