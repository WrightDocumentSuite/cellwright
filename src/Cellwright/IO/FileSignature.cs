namespace Cellwright.IO;

/// <summary>
/// Recognises what a stream actually holds before the package reader tries to open it, so that
/// the wrong kind of file produces a sentence a person can act on rather than a zip error.
/// </summary>
internal static class FileSignature
{
    private static ReadOnlySpan<byte> Zip => [0x50, 0x4B];

    private static ReadOnlySpan<byte> CompoundFile => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>
    /// Throws with a useful message when the stream is plainly not an xlsx package, and returns
    /// it to where it was otherwise.
    /// </summary>
    /// <exception cref="XlsxFormatException">The stream holds something other than a package.</exception>
    public static async ValueTask RejectNonPackageAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (!stream.CanSeek)
            return;

        long origin = stream.Position;
        byte[] head = new byte[8];
        int read = await stream.ReadAtLeastAsync(head, head.Length, throwOnEndOfStream: false, cancellationToken)
            .ConfigureAwait(false);
        stream.Position = origin;

        if (read >= 2 && head.AsSpan(0, 2).SequenceEqual(Zip))
            return;

        if (read >= 8 && head.AsSpan(0, 8).SequenceEqual(CompoundFile))
        {
            // Two very different files start these eight bytes: a legacy binary workbook, and a
            // modern package encrypted with a password to open — Excel does not encrypt the zip,
            // it puts the whole package inside a compound file. Telling them apart means reading
            // the container's directory, so the message names both rather than guessing.
            throw new XlsxFormatException(
                "This is a compound file, not an .xlsx package: either a legacy binary .xls workbook or an " +
                "xlsx encrypted with a password to open, which Excel wraps in the same container. Both need " +
                "the Cellwright.Xls package: install it and call Spreadsheet.LoadAsync, which recognizes " +
                "either and returns the same Workbook this method would.");
        }

        if (read >= 4 && IsLegacyRecordStream(head))
        {
            throw new XlsxFormatException(
                "This is a legacy binary Excel workbook from before Excel 5, not an .xlsx package. Reading it needs " +
                "the Cellwright.Xls package: install it and call XlsWorkbook.LoadAsync.");
        }

        throw new XlsxFormatException(
            "The stream does not hold an .xlsx package: it does not begin with the zip signature every one starts with.");
    }

    /// <summary>Whether the bytes open a bare BIFF record stream, as Excel 2 through 4 wrote.</summary>
    private static bool IsLegacyRecordStream(ReadOnlySpan<byte> head)
    {
        ushort type = (ushort)(head[0] | (head[1] << 8));
        ushort size = (ushort)(head[2] | (head[3] << 8));
        return type is 0x0009 or 0x0209 or 0x0409 && size is >= 4 and <= 16;
    }
}
