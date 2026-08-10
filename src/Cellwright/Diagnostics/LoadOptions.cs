namespace Cellwright;

/// <summary>
/// Options for <see cref="Workbook.LoadAsync(Stream, LoadOptions?, CancellationToken)"/>.
/// </summary>
public sealed record LoadOptions
{
    internal static readonly LoadOptions Default = new();

    private DocumentLoadBudget _budget = DocumentLoadBudget.Default;
    private PackageLimits? _legacyLimits;

    /// <summary>Keep the source stream open after loading. Ignored for file paths.</summary>
    public bool LeaveOpen { get; init; }

    /// <summary>
    /// Keep the package content the model does not represent — charts, pivot tables, images,
    /// macros, printer settings, conditional formatting — so that saving reproduces it.
    /// On by default: editing a real workbook should not cost it everything we do not model.
    /// Turn it off to load large files read-only without paying for the extra memory.
    /// </summary>
    public bool PreserveUnknownParts { get; init; } = true;

    /// <summary>
    /// Invoked as recoverable problems are found. All warnings are also collected into
    /// <see cref="Workbook.LoadDiagnostics"/> regardless of this callback.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; init; }

    /// <summary>
    /// Shared ceilings for source bytes, package expansion, XML, cells, media and embedded
    /// objects. A finite budget is applied by default.
    /// </summary>
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
    /// per-part and total-inflated ceilings in <see cref="Budget"/>; new code should configure
    /// <see cref="Budget"/> directly.
    /// </summary>
    public PackageLimits Limits
    {
        get => _legacyLimits ?? PackageLimits.From(_budget);
        init => _legacyLimits = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal DocumentLoadBudget EffectiveBudget => _legacyLimits?.ApplyTo(_budget) ?? _budget;
}

/// <summary>
/// Options for <see cref="Workbook.SaveAsync(Stream, SaveOptions?, CancellationToken)"/>.
/// </summary>
public sealed record SaveOptions
{
    internal static readonly SaveOptions Default = new();

    /// <summary>Write text as inline strings instead of a shared-string table.</summary>
    public bool UseInlineStrings { get; init; }

    /// <summary>Keep the destination stream open after saving. Ignored for file paths.</summary>
    public bool LeaveOpen { get; init; }

    /// <summary>
    /// Called for anything the saved file will not carry across — a drawing rewritten from the
    /// model, losing the shapes beside its pictures. Nothing here stops the save.
    /// </summary>
    public Action<WorkbookWarning>? OnWarning { get; init; }
}

/// <summary>
/// Everything recoverable that happened while loading a workbook. The library never fails
/// silently: repairs and substitutions surface here (and through <see cref="LoadOptions.OnWarning"/>).
/// </summary>
public sealed class LoadDiagnostics
{
    private readonly List<WorkbookWarning> _warnings = [];

    /// <summary>All warnings in occurrence order.</summary>
    public IReadOnlyList<WorkbookWarning> Warnings => _warnings;

    /// <summary>Whether anything noteworthy happened during the load.</summary>
    public bool HasWarnings => _warnings.Count > 0;

    internal void Add(WorkbookWarning warning) => _warnings.Add(warning);
}
