namespace Cellwright.Pdf;

/// <summary>
/// Ready-made values for <see cref="PdfExportOptions.RefuseWhen"/>. A mandatory report is better
/// not produced at all than produced with something missing.
/// </summary>
public static class PdfExportRefusal
{
    /// <summary>Refuses when an image or other content was left out of the document.</summary>
    public static Func<PdfExportWarning, bool> Omissions { get; } =
        static warning => warning.Kind is PdfExportWarningKind.ImageSkipped or PdfExportWarningKind.ContentSkipped;

    /// <summary>Refuses when content was left out or a font was drawn with another face.</summary>
    public static Func<PdfExportWarning, bool> OmissionsAndFontSubstitutions { get; } =
        static warning => Omissions(warning) || warning.Kind == PdfExportWarningKind.FontSubstituted;

    /// <summary>Refuses on any compromise at all, including layout that was only approximated.</summary>
    public static Func<PdfExportWarning, bool> AnyCompromise { get; } = static _ => true;
}

/// <summary>
/// Raised when a render made a compromise that <see cref="PdfExportOptions.RefuseWhen"/> does not
/// accept. No document is produced, and a file named by the caller is left as it was.
/// </summary>
public sealed class PdfExportRefusedException : Exception
{
    internal PdfExportRefusedException(IReadOnlyList<PdfExportWarning> refused, PdfExportDiagnostics diagnostics)
        : base($"The render was refused because of {refused.Count} compromise(s): {string.Join("; ", refused)}")
    {
        Refused = refused;
        Diagnostics = diagnostics;
    }

    /// <summary>The compromises the policy did not accept.</summary>
    public IReadOnlyList<PdfExportWarning> Refused { get; }

    /// <summary>Everything the render had to work around, accepted or not.</summary>
    public PdfExportDiagnostics Diagnostics { get; }
}
