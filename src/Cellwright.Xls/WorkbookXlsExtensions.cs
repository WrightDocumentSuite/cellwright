using Cellwright.Xls.Cfb;
using Cellwright.Xls.Crypto;

namespace Cellwright.Xls;

/// <summary>
/// Saving a workbook in the legacy binary format or behind a password, alongside the
/// <c>SaveAsync</c> the core already has for the modern one, and reading the macros it came
/// with.
/// </summary>
public static class WorkbookXlsExtensions
{
    /// <summary>
    /// Reads the macros the workbook was loaded with, or <see langword="null"/> when it has
    /// none. The project is not part of the model: it is one of the things the load keeps
    /// verbatim, so this needs <see cref="LoadOptions.PreserveUnknownParts"/> left on, and a
    /// workbook that came from an .xls has none regardless — that reader skips the project.
    /// </summary>
    /// <param name="workbook">Workbook loaded from a package.</param>
    /// <param name="options">Reading settings; only the fallback code page matters here.</param>
    public static VbaProject? GetVbaProject(this Workbook workbook, XlsLoadOptions? options = null) =>
        VbaProject.FromWorkbook(workbook, options);

    /// <summary>Writes the workbook to a file in the Excel 97-2003 binary format.</summary>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="path">Path of the .xls file to create.</param>
    /// <param name="options">Writing settings.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static ValueTask SaveXlsAsync(
        this Workbook workbook, string path, XlsSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        XlsWorkbook.SaveAsync(workbook, path, options, cancellationToken);

    /// <summary>Writes the workbook to a stream in the Excel 97-2003 binary format.</summary>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="destination">Stream to write to.</param>
    /// <param name="options">Writing settings.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static ValueTask SaveXlsAsync(
        this Workbook workbook, Stream destination, XlsSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        XlsWorkbook.SaveAsync(workbook, destination, options, cancellationToken);

    /// <summary>
    /// Writes the workbook as a modern package behind a password to open it.
    /// </summary>
    /// <remarks>
    /// The result is not an encrypted zip: Excel puts the whole package inside a compound file
    /// beside a stream describing how it was encrypted, so the file begins with the same eight
    /// bytes a 2003 workbook does. Read it back through <see cref="Spreadsheet.LoadAsync(string, XlsLoadOptions?, CancellationToken)"/>
    /// with the password; the core's own <c>Workbook.LoadAsync</c> sees a compound file and says so.
    /// <para>
    /// The package inside is an ordinary xlsx, written by the same saver as always, so nothing
    /// is lost by encrypting — the parts the model does not understand still travel through.
    /// </para>
    /// </remarks>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="path">Path of the file to create.</param>
    /// <param name="password">Password that will open it.</param>
    /// <param name="options">Writing settings for the package inside.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static async ValueTask SaveEncryptedAsync(
        this Workbook workbook, string path, string password,
        SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        FileStream stream = new(path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 1 << 16, useAsync: true);
        await using (stream.ConfigureAwait(false))
        {
            await SaveEncryptedAsync(workbook, stream, password, options, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <inheritdoc cref="SaveEncryptedAsync(Workbook, string, string, SaveOptions?, CancellationToken)"/>
    /// <param name="workbook">Workbook to write.</param>
    /// <param name="destination">Stream to write to.</param>
    /// <param name="password">Password that will open it.</param>
    /// <param name="options">Writing settings for the package inside.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    public static async ValueTask SaveEncryptedAsync(
        this Workbook workbook, Stream destination, string password,
        SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(destination);

        // An empty password would produce a file that opens without one being asked for, which
        // is not encryption failing loudly but encryption failing quietly.
        ArgumentException.ThrowIfNullOrEmpty(password);

        // The package has to exist whole before it can be encrypted: the integrity value covers
        // the finished ciphertext, so there is nothing to stream past this point.
        using var package = new MemoryStream();
        await workbook.SaveAsync(package, (options ?? new SaveOptions()) with { LeaveOpen = true }, cancellationToken)
            .ConfigureAwait(false);

        await CompoundFileWriter.WriteAsync(
            destination,
            OoxmlEncryption.Encrypt(package.GetBuffer().AsSpan(0, (int)package.Length), password),
            cancellationToken).ConfigureAwait(false);
    }
}
