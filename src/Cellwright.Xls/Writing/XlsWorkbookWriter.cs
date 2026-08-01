using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes a workbook as a BIFF8 record stream, ready to be wrapped in a compound file.
/// </summary>
/// <remarks>
/// The order is forced by the format: the sheets have to be written before the global substream
/// can state where they start, and before the shared string table knows what is in it. So the
/// sheets go into a buffer of their own first, and the two streams are joined once both are
/// complete and the sheet offsets can be filled in.
/// </remarks>
internal static class XlsWorkbookWriter
{
    /// <summary>Writes the workbook and returns the bytes of the <c>Workbook</c> stream.</summary>
    public static byte[] Write(Workbook workbook, XlsSaveOptions options)
    {
        void Warn(WorkbookWarning warning) => options.OnWarning?.Invoke(warning);

        if (workbook.Sheets.Count == 0)
            throw new InvalidOperationException("A workbook must have at least one sheet to be written.");

        StyleTable styles = StyleTable.Build(workbook.StylePool, Warn);
        var context = new XlsWriteContext(workbook, Warn)
        {
            XfByStyle = styles.XfByStyle,
            TemporalByXf = styles.TemporalByXf,
            StyleTable = styles,
        };

        // The sheets are written first because the shared string table they fill has to be
        // complete before the global substream that carries it is emitted.
        using var sheetStream = new BiffWriter(1 << 18);
        var sheetStarts = new List<int>(workbook.Sheets.Count);
        for (int index = 0; index < workbook.Sheets.Count; index++)
        {
            sheetStarts.Add(sheetStream.Length);
            new SheetWriter(context, workbook.Sheets[index], index).Write(sheetStream);
        }

        using var globalStream = new BiffWriter(1 << 16);
        var globals = new GlobalsWriter(context, styles);
        globals.Write(globalStream);

        int globalLength = globalStream.Length;
        for (int i = 0; i < sheetStarts.Count && i < globals.SheetOffsetFields.Count; i++)
            globalStream.PatchUInt32(globals.SheetOffsetFields[i], (uint)(globalLength + sheetStarts[i]));

        var result = new byte[globalLength + sheetStream.Length];
        globalStream.Written.Span.CopyTo(result);
        sheetStream.Written.Span.CopyTo(result.AsSpan(globalLength));
        return result;
    }
}
