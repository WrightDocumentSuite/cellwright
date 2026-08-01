using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the print settings of a sheet. They belong to two separate blocks of the substream:
/// the flags and the page breaks go with the sheet globals, the paper and the margins into the
/// page-setup block after them, and both have a fixed order the reader relies on.
/// </summary>
internal static class PrintSettingsWriter
{
    /// <summary>Header and footer margins Excel writes when the sheet states none.</summary>
    private const double DefaultHeaderMargin = 0.5;

    /// <summary>
    /// Writes the part of the settings that belongs to the sheet globals: what is printed
    /// besides the cells, the outline defaults, and the manual page breaks.
    /// </summary>
    public static void WriteGlobals(BiffWriter writer, PageSetup setup)
    {
        WriteFlag(writer, BiffRecordType.PrintRowCol, setup.PrintHeadings);
        WriteFlag(writer, BiffRecordType.PrintGrid, setup.PrintGridLines);
        WriteFlag(writer, BiffRecordType.GridSet, true);

        int flags = 0x04C1;             // Summary rows below, summary columns right, styles applied.
        if (!setup.AutoPageBreaks)
            flags &= ~0x0001;
        if (setup.FitToPage)
            flags |= 0x0100;
        WriteUInt16Record(writer, BiffRecordType.WsBool, flags);

        WriteBreaks(writer, BiffRecordType.HorizontalPageBreaks, setup.RowBreaks, span: 255);
        WriteBreaks(writer, BiffRecordType.VerticalPageBreaks, setup.ColumnBreaks, span: 65535);
    }

    /// <summary>Writes the page-setup block: headers, centering, margins and the paper itself.</summary>
    public static void WritePageSetup(BiffWriter writer, PageSetup setup)
    {
        WriteText(writer, BiffRecordType.Header, setup.Header?.Default);
        WriteText(writer, BiffRecordType.Footer, setup.Footer?.Default);
        WriteFlag(writer, BiffRecordType.HCenter, setup.CenterHorizontally);
        WriteFlag(writer, BiffRecordType.VCenter, setup.CenterVertically);

        if (setup.Margins is { } margins)
        {
            WriteMargin(writer, BiffRecordType.LeftMargin, margins.Left);
            WriteMargin(writer, BiffRecordType.RightMargin, margins.Right);
            WriteMargin(writer, BiffRecordType.TopMargin, margins.Top);
            WriteMargin(writer, BiffRecordType.BottomMargin, margins.Bottom);
        }

        WriteSetup(writer, setup);
        WriteHeaderFooterExtension(writer, setup);
    }

    private static void WriteSetup(BiffWriter writer, PageSetup setup)
    {
        int flags = 0;
        if (setup.Order == PageOrder.OverThenDown)
            flags |= 0x0001;
        if (setup.Orientation != PageOrientation.Landscape)
            flags |= 0x0002;                            // Portrait, which is also the default.
        if (setup.Orientation == PageOrientation.Default)
            flags |= 0x0040;                            // The orientation was never chosen.
        if (setup.BlackAndWhite)
            flags |= 0x0008;
        if (setup.Draft)
            flags |= 0x0010;
        if (setup.FirstPageNumber > 0)
            flags |= 0x0080;

        int start = writer.BeginRecord(BiffRecordType.Setup);
        writer.AppendUInt16(setup.PaperSize);
        writer.AppendUInt16(Math.Clamp(setup.Scale, 10, 400));
        writer.AppendUInt16(setup.FirstPageNumber > 0 ? setup.FirstPageNumber : 1);
        writer.AppendUInt16(Math.Clamp(setup.FitToWidth ?? 1, 0, 32767));
        writer.AppendUInt16(Math.Clamp(setup.FitToHeight ?? 1, 0, 32767));
        writer.AppendUInt16(flags);
        writer.AppendUInt16(600);                       // Print resolution, in dots per inch.
        writer.AppendUInt16(600);
        writer.AppendDouble(setup.Margins?.Header ?? DefaultHeaderMargin);
        writer.AppendDouble(setup.Margins?.Footer ?? DefaultHeaderMargin);
        writer.AppendUInt16(1);                         // Copies.
        writer.EndRecord(start);
    }

    /// <summary>
    /// Writes the record that carries the even and first-page headers, which the two older
    /// records have no room for. It is only written when one of them is actually in use.
    /// </summary>
    private static void WriteHeaderFooterExtension(BiffWriter writer, PageSetup setup)
    {
        if (!setup.DifferentOddEven && !setup.DifferentFirstPage)
            return;

        string? evenHeader = setup.DifferentOddEven ? setup.Header?.Even : null;
        string? evenFooter = setup.DifferentOddEven ? setup.Footer?.Even : null;
        string? firstHeader = setup.DifferentFirstPage ? setup.Header?.First : null;
        string? firstFooter = setup.DifferentFirstPage ? setup.Footer?.First : null;

        int flags = 0;
        if (setup.DifferentOddEven)
            flags |= 0x0001;
        if (setup.DifferentFirstPage)
            flags |= 0x0002;
        if (setup.ScaleHeaderWithDocument)
            flags |= 0x0004;
        if (setup.AlignHeaderWithMargins)
            flags |= 0x0008;

        int start = writer.BeginRecord(BiffRecordType.HeaderFooter);
        writer.AppendUInt16(BiffRecordType.HeaderFooter);   // Future-record header: the type,
        writer.AppendUInt16(0);                             // its flags, and a reference
        writer.AppendZeros(8);                              // to the range it applies to.
        writer.AppendZeros(16);                             // Sheet view this belongs to: none.
        writer.AppendUInt16(flags);
        writer.AppendUInt16(evenHeader?.Length ?? 0);
        writer.AppendUInt16(evenFooter?.Length ?? 0);
        writer.AppendUInt16(firstHeader?.Length ?? 0);
        writer.AppendUInt16(firstFooter?.Length ?? 0);

        foreach (string? text in new[] { evenHeader, evenFooter, firstHeader, firstFooter })
        {
            if (!string.IsNullOrEmpty(text))
                writer.AppendUnicodeString(text);
        }

        writer.EndRecord(start);
    }

    private static void WriteBreaks(BiffWriter writer, ushort type, IReadOnlyList<int> breaks, int span)
    {
        // One record holds at most 1026 breaks, and a sheet that reaches the limit has bigger
        // problems than the ones past it going missing.
        if (breaks.Count == 0)
            return;

        int count = Math.Min(breaks.Count, 1026);
        int start = writer.BeginRecord(type);
        writer.AppendUInt16(count);
        for (int i = 0; i < count; i++)
        {
            writer.AppendUInt16(breaks[i] - 1);
            writer.AppendUInt16(0);
            writer.AppendUInt16(span);
        }

        writer.EndRecord(start);
    }

    private static void WriteText(BiffWriter writer, ushort type, string? text)
    {
        int start = writer.BeginRecord(type);
        if (!string.IsNullOrEmpty(text))
            writer.AppendUnicodeString(text.Length > 255 ? text[..255] : text);
        writer.EndRecord(start);
    }

    private static void WriteMargin(BiffWriter writer, ushort type, double inches)
    {
        int start = writer.BeginRecord(type);
        writer.AppendDouble(inches);
        writer.EndRecord(start);
    }

    private static void WriteFlag(BiffWriter writer, ushort type, bool value) =>
        WriteUInt16Record(writer, type, value ? 1 : 0);

    private static void WriteUInt16Record(BiffWriter writer, ushort type, int value)
    {
        int start = writer.BeginRecord(type);
        writer.AppendUInt16(value);
        writer.EndRecord(start);
    }
}
