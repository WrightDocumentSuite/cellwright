namespace Cellwright;

/// <summary>
/// Options controlling how <see cref="XlsxReader"/> consumes a workbook.
/// </summary>
public sealed record XlsxReaderOptions
{
    internal static readonly XlsxReaderOptions Default = new();

    private DocumentLoadBudget _budget = DocumentLoadBudget.Default;
    private PackageLimits? _legacyLimits;

    /// <summary>Keep the source stream open after the reader is disposed. Ignored for file paths.</summary>
    public bool LeaveOpen { get; init; }

    /// <summary>Shared ceilings for source bytes, package expansion, XML, cells and resources.</summary>
    public DocumentLoadBudget Budget
    {
        get => EffectiveBudget;
        init
        {
            _budget = value ?? throw new ArgumentNullException(nameof(value));
            _legacyLimits = null;
        }
    }

    /// <summary>
    /// Compatibility view of the former package-only limits. Setting it updates input,
    /// per-part and total-inflated ceilings in <see cref="Budget"/>.
    /// </summary>
    public PackageLimits Limits
    {
        get => _legacyLimits ?? PackageLimits.From(_budget);
        init => _legacyLimits = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal DocumentLoadBudget EffectiveBudget => _legacyLimits?.ApplyTo(_budget) ?? _budget;

    /// <summary>
    /// Invoked for recoverable oddities found while reading (unknown cell types, malformed values,
    /// out-of-range shared string indexes). The reader substitutes a safe value and continues.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; init; }
}

/// <summary>
/// Identity of one worksheet discovered in a workbook.
/// </summary>
/// <param name="Index">Zero-based position in the workbook.</param>
/// <param name="Name">Sheet name.</param>
/// <param name="Visibility">How visible the sheet is in the Excel UI.</param>
/// <param name="SheetId">
/// The identifier the workbook part gave the sheet. Revisions and extension lists point at
/// sheets by this number, so it survives a round trip rather than being reassigned.
/// </param>
public sealed record SheetInfo(int Index, string Name, SheetVisibility Visibility, int SheetId = 0)
{
    /// <summary>Whether the sheet is hidden at all, either way.</summary>
    public bool IsHidden => Visibility != SheetVisibility.Visible;
}

/// <summary>
/// A defined name as declared in the workbook part.
/// </summary>
/// <param name="Name">Name used in formulas.</param>
/// <param name="RefersTo">Target formula, e.g. <c>Sheet1!$A$1:$B$5</c>.</param>
/// <param name="LocalSheetIndex">Zero-based worksheet index for sheet-scoped names, otherwise <see langword="null"/>.</param>
/// <param name="IsHidden">Whether Excel hides the name from the name manager.</param>
public sealed record DefinedNameInfo(string Name, string RefersTo, int? LocalSheetIndex, bool IsHidden = false)
{
    /// <summary>
    /// Attributes the model does not represent: the comment shown in the name manager, the
    /// shortcut key, and the flags that make a name a macro rather than a reference.
    /// </summary>
    internal List<KeyValuePair<string, string>>? Attributes { get; init; }
}
