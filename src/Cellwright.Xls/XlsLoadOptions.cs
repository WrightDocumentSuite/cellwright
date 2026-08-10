namespace Cellwright.Xls;

/// <summary>
/// Settings for reading a legacy workbook.
/// </summary>
public sealed class XlsLoadOptions
{
    private DocumentLoadBudget _budget = DocumentLoadBudget.Default;
    private PackageLimits? _legacyLimits;

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

    /// <summary>Shared ceilings for source bytes, container/package expansion, cells and resources.</summary>
    public DocumentLoadBudget Budget
    {
        get => EffectiveBudget;
        set
        {
            _budget = value ?? throw new ArgumentNullException(nameof(value));
            _legacyLimits = null;
        }
    }

    /// <summary>
    /// Compatibility view of the former package-only limits. Setting it updates input,
    /// per-stream and total-inflated ceilings in <see cref="Budget"/>.
    /// </summary>
    public PackageLimits Limits
    {
        get => _legacyLimits ?? PackageLimits.From(_budget);
        set => _legacyLimits = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal DocumentLoadBudget EffectiveBudget => _legacyLimits?.ApplyTo(_budget) ?? _budget;

    /// <summary>Whether to leave the source stream open after reading.</summary>
    public bool LeaveOpen { get; set; }

    /// <summary>Called for every recoverable problem, as it is found.</summary>
    public Action<WorkbookWarning>? OnWarning { get; set; }
}
