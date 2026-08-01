using Cellwright.IO;
using Cellwright.Xls.Cfb;
using Cellwright.Xls.Crypto;
using Cellwright.Xls.Reading;
using Cellwright.Xls.Xlsb;

namespace Cellwright.Xls;

/// <summary>
/// Opens a spreadsheet without being told which format it is in. Both the modern package and
/// every generation of the binary format land in the same <see cref="Workbook"/>, so code that
/// takes a file from a user does not have to branch on the extension — which is unreliable
/// anyway, since plenty of files named <c>.xls</c> are anything but.
/// </summary>
public static class Spreadsheet
{
    /// <summary>Opens a workbook from a file, choosing the reader by what the file holds.</summary>
    /// <param name="path">Path of the workbook.</param>
    /// <param name="options">Settings for the legacy reader, and the password of an encrypted file of either format.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    public static async ValueTask<Workbook> LoadAsync(
        string path, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        await using FileStream stream = new(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 16, useAsync: true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Opens a workbook from a stream, choosing the reader by what the stream holds.</summary>
    /// <param name="stream">Stream positioned at the start of the workbook.</param>
    /// <param name="options">Settings for the legacy reader, and the password of an encrypted file of either format.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <exception cref="XlsFormatException">The stream holds neither format, or no password opens it.</exception>
    public static async ValueTask<Workbook> LoadAsync(
        Stream stream, XlsLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        options ??= new XlsLoadOptions();

        Format format = await DetectAsync(stream, cancellationToken).ConfigureAwait(false);
        if (format == Format.Package)
            return await LoadPackageAsync(stream, options, options.LeaveOpen, cancellationToken).ConfigureAwait(false);

        byte[] bytes = await XlsWorkbook.ReadAllAsync(stream, options.LeaveOpen, options.Limits, cancellationToken).ConfigureAwait(false);
        if (TryDecryptPackage(bytes, options) is not { } decrypted)
            return XlsWorkbookReader.Read(bytes, options);

        using var package = new MemoryStream(decrypted, writable: false);
        return await LoadPackageAsync(package, options, leaveOpen: true, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a zip package, which holds either the XML format or the binary one. Which it is
    /// shows in the name of the part the <c>officeDocument</c> relationship points at: a
    /// <c>.bin</c> workbook is records rather than markup, and everything hanging off it is too.
    /// </summary>
    private static async ValueTask<Workbook> LoadPackageAsync(
        Stream stream, XlsLoadOptions options, bool leaveOpen, CancellationToken cancellationToken)
    {
        long origin = stream.CanSeek ? stream.Position : 0;
        OpcPackage package = await OpcPackage
            .OpenReadAsync(stream, leaveOpen: true, cancellationToken, options.Limits)
            .ConfigureAwait(false);

        string? binary;
        try
        {
            binary = await FindBinaryWorkbookAsync(package, cancellationToken).ConfigureAwait(false);
            if (binary is not null)
                return await XlsbWorkbookReader.ReadAsync(package, binary, options, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            await package.DisposeAsync().ConfigureAwait(false);
        }

        // The XML reader opens the package itself, so it gets the stream back where it started.
        if (stream.CanSeek)
            stream.Position = origin;

        var loadOptions = new LoadOptions
        {
            LeaveOpen = leaveOpen,
            OnWarning = options.OnWarning,
            Limits = options.Limits,
        };
        return await Workbook.LoadAsync(stream, loadOptions, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// The workbook part of a binary package, or <see langword="null"/> when the package holds
    /// the XML format and the ordinary reader should have it. Which one it is shows in the name
    /// of the part the <c>officeDocument</c> relationship points at: a <c>.bin</c> workbook is
    /// records rather than markup, and every part hanging off it is too.
    /// </summary>
    private static async ValueTask<string?> FindBinaryWorkbookAsync(
        OpcPackage package, CancellationToken cancellationToken)
    {
        IReadOnlyList<OpcRelationship> relationships =
            await package.GetRelationshipsAsync("/", cancellationToken).ConfigureAwait(false);

        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Type != XlsxSchema.RelOfficeDocument || relationship.IsExternal)
                continue;

            string target = OpcPath.Resolve("/", relationship.Target);
            return target.EndsWith(".bin", StringComparison.OrdinalIgnoreCase) ? target : null;
        }

        return null;
    }

    /// <summary>
    /// Unwraps a package that was encrypted with a password to open, or returns
    /// <see langword="null"/> when the container holds an ordinary binary workbook instead.
    /// </summary>
    private static byte[]? TryDecryptPackage(byte[] bytes, XlsLoadOptions options)
    {
        if (!CompoundFile.HasSignature(bytes))
            return null;

        CompoundFile container = CompoundFile.Parse(bytes, options.Limits);
        if (!OoxmlEncryption.IsEncryptedPackage(container))
            return null;

        byte[] package = OoxmlEncryption.Decrypt(container, options.Password, options.OnWarning);
        options.OnWarning?.Invoke(new WorkbookWarning(
            "EncryptedWorkbook",
            "The package was encrypted with a password to open. It was decrypted for reading; saving writes it back in the clear."));
        return package;
    }

    private enum Format
    {
        Package,
        Legacy,
    }

    /// <summary>
    /// Reads the first bytes to see which format this is, then puts the stream back. A stream
    /// that cannot seek is read as legacy, since that reader buffers the whole thing anyway.
    /// </summary>
    private static async ValueTask<Format> DetectAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (!stream.CanSeek)
            return Format.Legacy;

        long origin = stream.Position;
        byte[] head = new byte[8];
        int read = await stream
            .ReadAtLeastAsync(head, head.Length, throwOnEndOfStream: false, cancellationToken)
            .ConfigureAwait(false);
        stream.Position = origin;

        return read >= 2 && head[0] == 0x50 && head[1] == 0x4B ? Format.Package : Format.Legacy;
    }
}
