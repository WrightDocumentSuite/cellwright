namespace Cellwright;

/// <summary>
/// What stays possible on a protected sheet. Excel stores these as a list of things the
/// protection does <em>not</em> cover, which is why the defaults here are the permissive ones:
/// turning protection on and changing nothing else locks the cells and leaves everything else.
/// </summary>
public sealed class SheetProtection
{
    /// <summary>Whether the sheet is protected at all.</summary>
    public bool IsEnabled { get; set; }

    /// <summary>Whether locked cells can still be selected.</summary>
    public bool AllowSelectLockedCells { get; set; } = true;

    /// <summary>Whether unlocked cells can still be selected.</summary>
    public bool AllowSelectUnlockedCells { get; set; } = true;

    /// <summary>Whether cell formatting can be changed.</summary>
    public bool AllowFormatCells { get; set; }

    /// <summary>Whether column width and visibility can be changed.</summary>
    public bool AllowFormatColumns { get; set; }

    /// <summary>Whether row height and visibility can be changed.</summary>
    public bool AllowFormatRows { get; set; }

    /// <summary>Whether columns can be inserted.</summary>
    public bool AllowInsertColumns { get; set; }

    /// <summary>Whether rows can be inserted.</summary>
    public bool AllowInsertRows { get; set; }

    /// <summary>Whether hyperlinks can be inserted.</summary>
    public bool AllowInsertHyperlinks { get; set; }

    /// <summary>Whether columns can be deleted.</summary>
    public bool AllowDeleteColumns { get; set; }

    /// <summary>Whether rows can be deleted.</summary>
    public bool AllowDeleteRows { get; set; }

    /// <summary>Whether ranges can be sorted.</summary>
    public bool AllowSort { get; set; }

    /// <summary>Whether the auto filter dropdowns still work.</summary>
    public bool AllowAutoFilter { get; set; }

    /// <summary>Whether pivot tables can be used.</summary>
    public bool AllowPivotTables { get; set; }

    /// <summary>Whether drawing objects can be edited.</summary>
    public bool AllowEditObjects { get; set; }

    /// <summary>Whether scenarios can be edited.</summary>
    public bool AllowEditScenarios { get; set; }

    /// <summary>Password verifier, or <see langword="null"/> when protection needs no password.</summary>
    public PasswordHash? Password { get; set; }

    /// <summary>Ranges left editable inside a protected sheet.</summary>
    public IList<ProtectedRange> EditableRanges { get; } = new List<ProtectedRange>();

    /// <summary>Turns protection on, optionally behind a password.</summary>
    public void Protect(string? password = null)
    {
        IsEnabled = true;
        Password = password is null ? null : PasswordHash.Create(password);
    }

    /// <summary>Turns protection off and forgets the password.</summary>
    public void Unprotect()
    {
        IsEnabled = false;
        Password = null;
    }
}

/// <summary>A range of a protected sheet that stays editable, optionally behind its own password.</summary>
public sealed class ProtectedRange
{
    /// <summary>Name shown in Excel's list of allowed ranges.</summary>
    public required string Name { get; init; }

    /// <summary>Cells the exception covers.</summary>
    public required IReadOnlyList<RangeRef> Ranges { get; set; }

    /// <summary>Password guarding the exception, or <see langword="null"/> when it is open.</summary>
    public PasswordHash? Password { get; init; }

    /// <summary>Creates an exception over an A1 range list such as <c>B2:D10</c>.</summary>
    public static ProtectedRange For(string name, string ranges, string? password = null) => new()
    {
        Name = name,
        Ranges = DataValidation.ParseRanges(ranges),
        Password = password is null ? null : PasswordHash.Create(password),
    };

    /// <inheritdoc/>
    public override string ToString() => $"{Name} over {string.Join(' ', Ranges)}";
}

/// <summary>
/// What a workbook allows once it is protected: reordering sheets and moving the window.
/// </summary>
public sealed class WorkbookProtection
{
    /// <summary>Whether sheets may not be added, removed, renamed or reordered.</summary>
    public bool LockStructure { get; set; }

    /// <summary>Whether the workbook window may not be moved or resized.</summary>
    public bool LockWindows { get; set; }

    /// <summary>Password verifier, or <see langword="null"/> when protection needs no password.</summary>
    public PasswordHash? Password { get; set; }

    /// <summary>Whether anything is protected.</summary>
    public bool IsEnabled => LockStructure || LockWindows;

    /// <summary>Locks the sheet list, optionally behind a password.</summary>
    public void Protect(string? password = null)
    {
        LockStructure = true;
        Password = password is null ? null : PasswordHash.Create(password);
    }

    /// <summary>Removes the protection and forgets the password.</summary>
    public void Unprotect()
    {
        LockStructure = false;
        LockWindows = false;
        Password = null;
    }
}
