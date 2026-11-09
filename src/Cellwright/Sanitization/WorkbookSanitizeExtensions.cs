namespace Cellwright.Sanitization;

/// <summary>Saves a copy of a workbook with chosen categories of content taken out.</summary>
public static class WorkbookSanitizeExtensions
{
    /// <summary>Writes a copy without the given categories. The workbook itself is not changed.</summary>
    /// <param name="workbook">The workbook to copy.</param>
    /// <param name="stream">Where the copy goes; it is left open.</param>
    /// <param name="remove">What to leave out.</param>
    /// <param name="cancellationToken">Cancels the save.</param>
    /// <returns>What was left out and what was kept.</returns>
    public static async ValueTask<SanitizeReport> SaveSanitizedAsync(
        this Workbook workbook,
        Stream stream,
        ContentCategories remove,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(stream);

        using var buffer = new MemoryStream();
        await workbook.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, cancellationToken).ConfigureAwait(false);

        (byte[] package, SanitizeReport report) = new XlsxContentSanitizer(remove).Run(buffer.ToArray());
        await stream.WriteAsync(package, cancellationToken).ConfigureAwait(false);
        return report;
    }

    /// <summary>Writes a copy without the given categories to a file, replacing it if it exists.</summary>
    /// <param name="workbook">The workbook to copy.</param>
    /// <param name="path">Destination path. A copy without macros should not keep an <c>.xlsm</c> name.</param>
    /// <param name="remove">What to leave out.</param>
    /// <param name="cancellationToken">Cancels the save.</param>
    /// <returns>What was left out and what was kept.</returns>
    public static async ValueTask<SanitizeReport> SaveSanitizedAsync(
        this Workbook workbook,
        string path,
        ContentCategories remove,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);

        using var buffer = new MemoryStream();
        SanitizeReport report = await workbook.SaveSanitizedAsync(buffer, remove, cancellationToken).ConfigureAwait(false);
        await File.WriteAllBytesAsync(path, buffer.ToArray(), cancellationToken).ConfigureAwait(false);
        return report;
    }
}
