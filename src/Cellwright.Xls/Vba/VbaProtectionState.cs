namespace Cellwright.Xls;

/// <summary>
/// Who closed a VBA project, and whether it shows at all. The file records three separate
/// decisions rather than one flag: a person can lock the project in the editor, the application
/// hosting the document can lock it, and the editor itself can.
/// </summary>
/// <remarks>
/// None of this hides the source. The state governs what the VBA editor offers to show; the
/// code is in the file either way, and comes out of <see cref="VbaProject.Modules"/> the same.
/// </remarks>
/// <param name="LockedByUser">The person editing the project asked for it to be protected.</param>
/// <param name="LockedByHost">The application hosting the document asked for it.</param>
/// <param name="LockedByEditor">The VBA editor asked for it.</param>
/// <param name="IsVisible">
/// Whether the project appears in the editor at all. A hidden project is also locked by the
/// editor, which is what makes hiding it stick.
/// </param>
public readonly record struct VbaProtectionState(
    bool LockedByUser, bool LockedByHost, bool LockedByEditor, bool IsVisible)
{
    /// <summary>Nothing protected and nothing hidden — what a project carries by default.</summary>
    public static VbaProtectionState Open { get; } = new(false, false, false, true);

    /// <summary>Whether anybody asked for the project to be closed.</summary>
    public bool IsLocked => LockedByUser || LockedByHost || LockedByEditor;
}
