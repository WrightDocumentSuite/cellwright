using System.Buffers.Binary;
using Cellwright.Model;
using Cellwright.Styles;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Cfb;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads a legacy workbook end to end: unwraps the container, decrypts if it has to, reads the
/// global substream, then each sheet, and hands back the same <see cref="Workbook"/> the xlsx
/// reader would have produced.
/// </summary>
internal static class XlsWorkbookReader
{
    /// <summary>Reads the bytes of a .xls file into a workbook.</summary>
    public static Workbook Read(
        byte[] bytes,
        XlsLoadOptions options,
        DocumentLoadBudgetState? loadBudget = null,
        CancellationToken cancellationToken = default,
        CompoundFile? container = null)
    {
        Workbook workbook = Workbook.Create();
        var diagnostics = new LoadDiagnostics();
        void Warn(WorkbookWarning warning)
        {
            diagnostics.Add(warning);
            options.OnWarning?.Invoke(warning);
        }

        loadBudget ??= new DocumentLoadBudgetState(options.EffectiveBudget);
        byte[] stream = ExtractWorkbookStream(
            bytes, options.EffectiveBudget, loadBudget, Warn, workbook.Properties, container);
        IRecordDecryptor? decryptor = FindDecryptor(stream, options, Warn);

        using var reader = new BiffReader(stream, decryptor);
        BiffVersion version = ReadOpeningBof(reader, Warn);

        // A BIFF2-4 file is one sheet with no global substream at all: the records that would
        // be global are mixed in with the cells.
        if (version <= BiffVersion.Biff4)
        {
            LegacySheetReader.Read(workbook, reader, version, options, Warn, loadBudget, cancellationToken);
            workbook.AttachDiagnostics(diagnostics);
            return workbook;
        }

        var globals = new GlobalsParser(
            reader, workbook.StylePool, Warn, loadBudget, cancellationToken).Parse(version);
        workbook.SetDateSystem(globals.DateSystem);

        var sheetNames = new List<string>(globals.Sheets.Count);
        foreach (XlsSheetEntry entry in globals.Sheets)
            sheetNames.Add(entry.Name);

        var decoder = new PtgDecoder(globals, sheetNames);
        ReadSheets(workbook, reader, globals, decoder, Warn, loadBudget, cancellationToken);
        DefinedNameLoader.Apply(workbook, globals, decoder, Warn);

        workbook.AttachDiagnostics(diagnostics);
        return workbook;
    }

    /// <summary>
    /// Prepares a file for row-by-row reading: unwraps the container, works out the encryption
    /// and reads the global substream, stopping short of the sheets themselves.
    /// </summary>
    /// <param name="bytes">Bytes of the .xls file.</param>
    /// <param name="options">Reading settings.</param>
    /// <param name="loadBudget">Shared resource counters for the import.</param>
    /// <param name="cancellationToken">Cancels record parsing.</param>
    /// <exception cref="XlsFormatException">The file is not a workbook this library can stream.</exception>
    public static (byte[] Stream, IRecordDecryptor? Decryptor, XlsGlobals Globals) OpenForStreaming(
        byte[] bytes,
        XlsLoadOptions options,
        DocumentLoadBudgetState? loadBudget = null,
        CancellationToken cancellationToken = default)
    {
        void Warn(WorkbookWarning warning) => options.OnWarning?.Invoke(warning);

        loadBudget ??= new DocumentLoadBudgetState(options.EffectiveBudget);
        byte[] stream = ExtractWorkbookStream(bytes, options.EffectiveBudget, loadBudget, Warn);
        IRecordDecryptor? decryptor = FindDecryptor(stream, options, Warn);

        using var reader = new BiffReader(stream, decryptor);
        BiffVersion version = ReadOpeningBof(reader, Warn);

        // A file older than Excel 5 has no sheet table to seek by, so there is nothing to
        // stream from: its records are one long sequence that has to be read as a whole.
        if (version <= BiffVersion.Biff4)
        {
            throw new XlsFormatException(
                "This workbook predates Excel 5 and has no sheet table, so its rows cannot be streamed. " +
                "Read it with XlsWorkbook.LoadAsync instead.");
        }

        // The pool is thrown away: streaming hands back values, and the styles only matter
        // here for telling a date from the number it is stored as.
        var globals = new GlobalsParser(
            reader, new StylePool(), Warn, loadBudget, cancellationToken).Parse(version);
        return (stream, decryptor, globals);
    }

    private static void ReadSheets(
        Workbook workbook,
        BiffReader reader,
        XlsGlobals globals,
        PtgDecoder decoder,
        Action<WorkbookWarning> warn,
        DocumentLoadBudgetState loadBudget,
        CancellationToken cancellationToken)
    {
        int globalsEnd = reader.Position;
        List<int>? substreamStarts = null;

        for (int index = 0; index < globals.Sheets.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XlsSheetEntry entry = globals.Sheets[index];
            if (entry.Kind != XlsSheetKind.Worksheet)
            {
                warn(new WorkbookWarning(
                    "SkippedSheet",
                    $"'{entry.Name}' is a {(entry.Kind == XlsSheetKind.Chart ? "chart" : "macro or dialog")} sheet, " +
                    "which this library does not model; it is not in the workbook that was loaded."));
                continue;
            }

            Worksheet sheet = WorkbookLoader.AddSheetSafe(workbook, entry.Name, warn);
            sheet.Visibility = entry.Visibility;

            int offset = entry.Offset;
            if (reader.PeekTypeAt(offset) is not BiffRecordType.Bof)
            {
                // Files exist whose sheet offsets are stale by a few bytes, usually because
                // something rewrote the globals without fixing them up. The substream is still
                // there; the nearest one is what the entry meant.
                substreamStarts ??= FindSubstreamStarts(reader, globalsEnd, cancellationToken);
                if (NearestStart(substreamStarts, offset) is not { } repaired)
                {
                    warn(new WorkbookWarning(
                        "BadSheetOffset",
                        $"The workbook says '{entry.Name}' starts at byte {offset}, where there is no sheet, and no " +
                        "sheet was found nearby; it was read as empty."));
                    continue;
                }

                warn(new WorkbookWarning(
                    "BadSheetOffset",
                    $"The workbook says '{entry.Name}' starts at byte {offset}; the sheet actually starts at {repaired}."));
                offset = repaired;
            }

            reader.Seek(offset);
            reader.Read();

            var sink = new WorksheetContentSink(
                sheet, styleMap: [], relationships: [], preserveMarkup: false, warn, loadBudget);
            new SheetParser(
                    reader,
                    globals,
                    decoder,
                    index,
                    NextSubstreamStart(reader, globals, offset),
                    warn,
                    loadBudget,
                    cancellationToken)
                .Parse(sink);
        }
    }

    /// <summary>
    /// Where the substream after this one begins, which bounds how far this sheet can run. A
    /// declared offset counts only if a substream really starts there: files exist whose sheet
    /// table is stale, and a stale offset must not cut a sheet short.
    /// </summary>
    private static int NextSubstreamStart(BiffReader reader, XlsGlobals globals, int start)
    {
        int nearest = int.MaxValue;
        foreach (XlsSheetEntry entry in globals.Sheets)
        {
            if (entry.Offset > start && entry.Offset < nearest &&
                reader.PeekTypeAt(entry.Offset) == BiffRecordType.Bof)
            {
                nearest = entry.Offset;
            }
        }

        return nearest;
    }

    /// <summary>
    /// Finds where each top-level substream begins. Nesting is tracked because an embedded
    /// chart opens a substream of its own inside the sheet that holds it.
    /// </summary>
    private static List<int> FindSubstreamStarts(
        BiffReader reader, int globalsEnd, CancellationToken cancellationToken)
    {
        int resume = reader.Position;
        var starts = new List<int>();
        int depth = 0;

        reader.Seek(globalsEnd);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            switch (reader.RecordType)
            {
                case BiffRecordType.Bof:
                    if (depth == 0)
                        starts.Add(reader.RecordOffset);
                    depth++;
                    break;

                case BiffRecordType.Eof:
                    depth = Math.Max(0, depth - 1);
                    break;

                default:
                    break;
            }
        }

        reader.Seek(resume);
        return starts;
    }

    /// <summary>Picks the substream start closest to where the workbook said one would be.</summary>
    private static int? NearestStart(List<int> starts, int declared)
    {
        int? best = null;
        int bestDistance = int.MaxValue;
        foreach (int start in starts)
        {
            int distance = Math.Abs(start - declared);
            if (distance >= bestDistance)
                continue;

            bestDistance = distance;
            best = start;
        }

        return best;
    }

    /// <summary>
    /// Unwraps the container, or takes the bytes as they are for a BIFF2-4 file. The document's
    /// properties live in streams beside the workbook rather than in it, so they are collected
    /// here, where the container is open, rather than by the record reader.
    /// </summary>
    private static byte[] ExtractWorkbookStream(
        byte[] bytes,
        DocumentLoadBudget budget,
        DocumentLoadBudgetState loadBudget,
        Action<WorkbookWarning> warn,
        WorkbookProperties? properties = null,
        CompoundFile? container = null)
    {
        if (!CompoundFile.HasSignature(bytes))
        {
            // Before Excel 5 a workbook was the record stream on its own, with no container.
            if (bytes.Length >= 4 && LooksLikeBof(bytes))
            {
                DocumentLoadBudgetState.Ensure(
                    nameof(DocumentLoadBudget.MaxPartBytes), budget.MaxPartBytes, bytes.LongLength);
                loadBudget.AddPackageParts(1);
                loadBudget.AddInflatedBytes(bytes.LongLength);
                return bytes;
            }

            throw new XlsFormatException(
                "The file is neither a compound file nor a bare BIFF stream, so it is not a legacy Excel workbook.");
        }

        container ??= CompoundFile.Parse(bytes, budget, loadBudget);
        if (Crypto.OoxmlEncryption.IsEncryptedPackage(container))
        {
            throw new XlsFormatException(
                "This compound file holds an .xlsx package encrypted with a password to open, not a legacy " +
                "BIFF workbook. Call Spreadsheet.LoadAsync with the password in XlsLoadOptions: it decrypts " +
                "the package and returns the same Workbook.");
        }

        if (properties is not null)
            DocumentProperties.Read(container, properties);

        // Excel 97 and later name it Workbook; Excel 5 and 95 named it Book. A file written by
        // both — which happens after a "save as" round trip — has each, and the newer wins.
        byte[]? stream = container.TryReadStream("Workbook") ?? container.TryReadStream("Book");
        if (stream is not null)
            return stream;

        foreach (string name in container.EntryNames)
        {
            if (name.Contains("Workbook", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Book", StringComparison.OrdinalIgnoreCase))
            {
                warn(new WorkbookWarning(
                    "UnusualStreamName",
                    $"The workbook stream is named '{name}' rather than 'Workbook'; it was read anyway."));
                return container.TryReadStream(name) ?? [];
            }
        }

        throw new XlsFormatException("The compound file holds no workbook stream.");
    }

    private static bool LooksLikeBof(ReadOnlySpan<byte> bytes)
    {
        ushort type = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        return type is 0x0009 or 0x0209 or 0x0409 or 0x0809;
    }

    /// <summary>
    /// Looks for a FilePass record near the start of the stream and builds a decryptor from it.
    /// It always sits in the first few records, right after the BOF.
    /// </summary>
    private static IRecordDecryptor? FindDecryptor(byte[] stream, XlsLoadOptions options, Action<WorkbookWarning> warn)
    {
        using var probe = new BiffReader(stream);
        for (int guard = 0; guard < 16 && probe.Read(); guard++)
        {
            if (probe.RecordType == BiffRecordType.Eof)
                break;
            if (probe.RecordType != BiffRecordType.FilePass)
                continue;

            if (XlsDecryptor.TryCreate(probe.Record.Span, options.Password, out XlsDecryptor? decryptor, out string? failure))
            {
                warn(new WorkbookWarning(
                    "EncryptedWorkbook",
                    "The workbook was encrypted; it was decrypted for reading and will be written back without encryption."));
                return decryptor;
            }

            throw new XlsFormatException(failure ?? "The workbook is encrypted and could not be decrypted.");
        }

        return null;
    }

    private static BiffVersion ReadOpeningBof(BiffReader reader, Action<WorkbookWarning> warn)
    {
        if (!reader.Read())
            throw new XlsFormatException("The workbook stream is empty.");

        ushort type = reader.RecordType;
        if (type is not (0x0009 or 0x0209 or 0x0409 or 0x0809))
            throw new XlsFormatException($"The workbook stream starts with record 0x{type:X4} rather than a BOF.");

        // Before BIFF5 the record's own type says which version wrote it; from BIFF5 on, the
        // type settled at 0x0809 and the version moved into the body.
        if (type != 0x0809)
        {
            return type switch
            {
                0x0009 => BiffVersion.Biff2,
                0x0209 => BiffVersion.Biff3,
                _ => BiffVersion.Biff4,
            };
        }

        ReadOnlySpan<byte> body = reader.Record.Span;
        if (body.Length < 4)
            return BiffVersion.Biff8;

        ushort version = BinaryPrimitives.ReadUInt16LittleEndian(body);
        return version switch
        {
            0x0500 => BiffVersion.Biff5,
            0x0600 => BiffVersion.Biff8,
            _ => Unrecognised(version, warn),
        };
    }

    private static BiffVersion Unrecognised(ushort version, Action<WorkbookWarning> warn)
    {
        warn(new WorkbookWarning(
            "UnknownBiffVersion",
            $"The workbook declares BIFF version 0x{version:X4}, which is not one this library knows; it was read as BIFF8."));
        return BiffVersion.Biff8;
    }
}
