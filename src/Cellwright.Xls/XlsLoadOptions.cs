namespace Cellwright.Xls;

/// <summary>
/// Settings for reading a legacy workbook.
/// </summary>
public sealed class XlsLoadOptions
{
    /// <summary>
    /// Password to open an encrypted workbook, binary or packaged: <see cref="Spreadsheet"/>
    /// uses it for an .xlsx locked with a password to open as well. Excel's own default password
    /// is tried first, so files that are encrypted without being protected open with nothing set here.
    /// </summary>
    public string? Password { get; set; }

    /// <summary>
    /// Code page for the byte strings of files older than Excel 97, when the file declares none
    /// or declares one this platform does not have. Defaults to Windows-1252.
    /// </summary>
    public int FallbackCodePage { get; set; } = 1252;

    /// <summary>
    /// Bounds the input held in memory and every stream extracted from an MS-CFB container.
    /// The defaults are shared with the OOXML reader; <see cref="PackageLimits.None"/> removes
    /// them for a file whose origin is trusted.
    /// </summary>
    public PackageLimits Limits { get; set; } = PackageLimits.Default;

    /// <summary>Whether to leave the source stream open after reading.</summary>
    public bool LeaveOpen { get; set; }

    /// <summary>Called for every recoverable problem, as it is found.</summary>
    public Action<WorkbookWarning>? OnWarning { get; set; }
}
