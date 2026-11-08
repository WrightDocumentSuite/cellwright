using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes one sheet's substream: the cell table and the records that describe the grid around
/// it.
/// </summary>
/// <remarks>
/// Cells are grouped into blocks of 32 rows with an index record after each, which is how Excel
/// finds a row without reading the whole sheet. The grouping is not optional — a reader that
/// trusts the index would go wrong without it — so the sheet's rows are gathered first and then
/// written block by block.
/// </remarks>
internal sealed class SheetWriter(XlsWriteContext context, Worksheet sheet, int sheetIndex)
{
    /// <summary>Rows per index block (MS-XLS 2.2.1).</summary>
    private const int RowsPerBlock = 32;

    /// <summary>The grid of the old format: 65536 rows by 256 columns.</summary>
    private const int MaxRow = 65536;

    private const int MaxColumn = 256;

    /// <summary>Excel's default column width in characters, which the model does not carry.</summary>
    private const double DefaultColumnWidth = 8.43;

    /// <summary>Excel's default row height in points.</summary>
    private const double DefaultRowHeight = 12.75;

    private readonly PtgEncoder _encoder = new(context);

    /// <summary>Writes the whole substream, in the order the format's grammar prescribes.</summary>
    public void Write(BiffWriter writer)
    {
        WriteBof(writer);

        List<int> rows = OccupiedRows(out int maxColumn);
        PrintSettingsWriter.WriteGlobals(writer, sheet.PageSetup);
        PrintSettingsWriter.WritePageSetup(writer, sheet.PageSetup);
        WriteProtection(writer);
        WriteColumns(writer);
        WriteDimensions(writer, rows, maxColumn);
        WriteCellTable(writer, rows);
        WriteNotes(writer);
        WriteWindow(writer);
        WriteMerges(writer);
        WriteConditionalFormats(writer);
        WriteHyperlinks(writer);
        WriteValidations(writer);

        writer.WriteRecord(BiffRecordType.Eof);
    }

    private static void WriteBof(BiffWriter writer)
    {
        Span<byte> body = stackalloc byte[16];
        body.Clear();
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, 0x0600);      // BIFF8.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[2..], 0x0010); // Worksheet.
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[4..], 0x0DBB);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[6..], 0x07CC);
        writer.WriteRecord(BiffRecordType.Bof, body);
    }

    /// <summary>
    /// Collects the rows that hold anything, clipping to the grid the old format has. A row past
    /// 65536 or a column past 256 simply cannot be written, and dropping it silently would be
    /// worse than saying so.
    /// </summary>
    private List<int> OccupiedRows(out int maxColumn)
    {
        var seen = new SortedSet<int>();
        maxColumn = 0;
        bool clipped = false;

        foreach ((int row, int column, _) in sheet.Values.EnumerateOccupied())
        {
            if (row > MaxRow || column > MaxColumn)
            {
                clipped = true;
                continue;
            }

            seen.Add(row);
            maxColumn = Math.Max(maxColumn, column);
        }

        foreach ((int row, int column, string? formula) in sheet.Formulas.EnumerateOccupied())
        {
            if (formula is null)
                continue;
            if (row > MaxRow || column > MaxColumn)
            {
                clipped = true;
                continue;
            }

            seen.Add(row);
            maxColumn = Math.Max(maxColumn, column);
        }

        foreach ((int row, int column, int style) in sheet.StyleIndexes.EnumerateOccupied())
        {
            if (style == 0)
                continue;
            if (row > MaxRow || column > MaxColumn)
            {
                clipped = true;
                continue;
            }

            seen.Add(row);
            maxColumn = Math.Max(maxColumn, column);
        }

        foreach (int row in sheet.RowTable.Keys)
        {
            if (row <= MaxRow)
                seen.Add(row);
        }

        if (clipped)
        {
            context.Warn(new WorkbookWarning(
                "GridTooLarge",
                $"'{sheet.Name}' has content beyond row {MaxRow} or column {MaxColumn}, which is as far as the " +
                ".xls grid goes; that content is not in the file."));
        }

        return [.. seen];
    }

    private void WriteDimensions(BiffWriter writer, List<int> rows, int maxColumn)
    {
        int firstRow = rows.Count > 0 ? rows[0] : 1;
        int lastRow = rows.Count > 0 ? rows[^1] : 1;

        int start = writer.BeginRecord(BiffRecordType.Dimensions);
        writer.AppendUInt32((uint)(firstRow - 1));
        writer.AppendUInt32((uint)lastRow);       // Exclusive.
        writer.AppendUInt16(0);
        writer.AppendUInt16(maxColumn);           // Exclusive.
        writer.AppendUInt16(0);                   // Reserved.
        writer.EndRecord(start);
    }

    private void WriteColumns(BiffWriter writer)
    {
        if (sheet.ColumnTable.Count == 0)
            return;

        // Neighbouring columns with the same settings share one record, which is how a file
        // that styles a hundred columns stays a few dozen bytes rather than a few thousand.
        foreach (int[] run in ColumnRuns())
        {
            int first = run[0];
            int last = run[1];
            ColumnProperties properties = sheet.ColumnTable[first];

            int start = writer.BeginRecord(BiffRecordType.ColInfo);
            writer.AppendUInt16(first - 1);
            writer.AppendUInt16(Math.Min(last, MaxColumn) - 1);
            writer.AppendUInt16((int)Math.Round((properties.Width ?? DefaultColumnWidth) * 256));
            writer.AppendUInt16(context.XfFor(properties.Style.Index));

            int flags = 0;
            if (properties.Hidden)
                flags |= 0x0001;
            if (properties.BestFit)
                flags |= 0x0004;
            flags |= (properties.OutlineLevel & 0x07) << 8;
            if (properties.Collapsed)
                flags |= 0x1000;
            writer.AppendUInt16(flags);
            writer.AppendUInt16(0);
            writer.EndRecord(start);
        }
    }

    /// <summary>Groups the styled columns into runs sharing identical settings.</summary>
    private IEnumerable<int[]> ColumnRuns()
    {
        List<int> columns = [.. sheet.ColumnTable.Keys.Where(static c => c <= MaxColumn).Order()];
        int index = 0;
        while (index < columns.Count)
        {
            int first = columns[index];
            ColumnProperties properties = sheet.ColumnTable[first];
            int last = first;
            while (index + 1 < columns.Count &&
                   columns[index + 1] == last + 1 &&
                   sheet.ColumnTable[columns[index + 1]] == properties)
            {
                last = columns[++index];
            }

            index++;
            yield return [first, last];
        }
    }

    /// <summary>
    /// Writes the cell table: the Row records of a block, then the cells of those rows, then
    /// the DBCell that indexes back to them.
    /// </summary>
    private void WriteCellTable(BiffWriter writer, List<int> rows)
    {
        for (int blockStart = 0; blockStart < rows.Count; blockStart += RowsPerBlock)
        {
            int blockEnd = Math.Min(blockStart + RowsPerBlock, rows.Count);
            int firstRowRecord = writer.Length;

            for (int i = blockStart; i < blockEnd; i++)
                WriteRowRecord(writer, rows[i]);

            var cellOffsets = new List<int>(blockEnd - blockStart);
            for (int i = blockStart; i < blockEnd; i++)
            {
                cellOffsets.Add(writer.Length);
                WriteRowCells(writer, rows[i]);
            }

            WriteDbCell(writer, firstRowRecord, cellOffsets);
        }
    }

    private void WriteRowRecord(BiffWriter writer, int row)
    {
        sheet.RowTable.TryGetValue(row, out RowProperties properties);

        int start = writer.BeginRecord(BiffRecordType.Row);
        writer.AppendUInt16(row - 1);
        writer.AppendUInt16(0);                                   // First column, filled below.
        writer.AppendUInt16(MaxColumn);                           // Last column, exclusive.
        writer.AppendUInt16(properties.Height is { } height
            ? (int)Math.Round(height * 20)
            : (int)Math.Round(DefaultRowHeight * 20) | 0x8000);
        writer.AppendUInt32(0);                                   // Reserved.

        uint flags = 0x00000100;                                  // Always 1, per the spec.
        flags |= (uint)(properties.OutlineLevel & 0x07);
        if (properties.Collapsed)
            flags |= 0x00000010;
        if (properties.Hidden)
            flags |= 0x00000020;
        if (properties.Height is not null)
            flags |= 0x00000040;
        if (properties.ThickTop)
            flags |= 0x10000000;
        if (properties.ThickBottom)
            flags |= 0x20000000;

        int styleIndex = properties.Style.Index;
        if (styleIndex != 0)
        {
            flags |= 0x00000080;                                  // The row has a format of its own.
            flags |= (uint)(context.XfFor(styleIndex) & 0x0FFF) << 16;
        }

        writer.AppendUInt32(flags);
        writer.EndRecord(start);
    }

    private void WriteRowCells(BiffWriter writer, int row)
    {
        int maxColumn = Math.Min(Math.Max(sheet.Values.MaxColumn, sheet.Formulas.MaxColumn), MaxColumn);
        maxColumn = Math.Max(maxColumn, Math.Min(sheet.StyleIndexes.MaxColumn, MaxColumn));

        for (int column = 1; column <= maxColumn; column++)
        {
            CellValue value = sheet.Values.Get(row, column);
            string? formula = sheet.Formulas.Get(row, column);
            int styleIndex = sheet.StyleIndexes.Get(row, column);
            if (value.IsBlank && formula is null && styleIndex == 0)
                continue;

            var cell = new CellRef(row, column);
            int xf = context.XfFor(styleIndex);
            if (formula is not null)
            {
                WriteFormulaCell(writer, cell, xf, formula, value);
                continue;
            }

            WriteValueCell(writer, cell, xf, value);
        }
    }

    private void WriteValueCell(BiffWriter writer, CellRef cell, int xf, in CellValue value)
    {
        switch (value.Kind)
        {
            case CellValueKind.Blank:
                WriteCellHeader(writer, BiffRecordType.Blank, cell, xf, bodySize: 0);
                return;

            case CellValueKind.Text:
                WriteLabelSst(writer, cell, xf, value.GetText());
                return;

            case CellValueKind.Boolean:
            case CellValueKind.Error:
                {
                    int start = writer.BeginRecord(BiffRecordType.BoolErr);
                    AppendCellHeader(writer, cell, xf);
                    if (value.Kind == CellValueKind.Boolean)
                    {
                        writer.AppendByte(value.GetBoolean() ? (byte)1 : (byte)0);
                        writer.AppendByte(0);
                    }
                    else
                    {
                        writer.AppendByte(BiffValues.EncodeError(value.GetError()));
                        writer.AppendByte(1);
                    }

                    writer.EndRecord(start);
                    return;
                }

            default:
                WriteNumericCell(writer, cell, xf, value.ToSerialNumber(context.Workbook.DateSystem));
                return;
        }
    }

    /// <summary>
    /// Writes a number as RK when it fits in four bytes, and as a full double when it does not.
    /// Most numbers in a sheet fit, so this halves the size of a numeric column.
    /// </summary>
    private static void WriteNumericCell(BiffWriter writer, CellRef cell, int xf, double number)
    {
        if (BiffValues.TryEncodeRk(number, out uint rk))
        {
            int rkStart = writer.BeginRecord(BiffRecordType.Rk);
            AppendCellHeader(writer, cell, xf);
            writer.AppendUInt32(rk);
            writer.EndRecord(rkStart);
            return;
        }

        int start = writer.BeginRecord(BiffRecordType.Number);
        AppendCellHeader(writer, cell, xf);
        writer.AppendDouble(number);
        writer.EndRecord(start);
    }

    private void WriteLabelSst(BiffWriter writer, CellRef cell, int xf, string text)
    {
        bool wasTruncated = false;
        if (text.Length > SheetLimits.MaxCellTextLength)
        {
            context.Warn(new WorkbookWarning(
                "TextTooLong",
                $"A cell of '{sheet.Name}' holds {text.Length} characters; the .xls format allows " +
                $"{SheetLimits.MaxCellTextLength}, so the text was cut short.",
                cell));
            text = text[..SheetLimits.MaxCellTextLength];
            wasTruncated = true;
        }

        int sstIndex;
        if (!wasTruncated &&
            sheet.RichStringTable.TryGetValue(cell, out IO.RichStringMarkup stored) &&
            string.Equals(stored.Text, text, StringComparison.Ordinal))
        {
            RichText rich = Streaming.RichTextMarkup.Parse(stored.Markup);
            sstIndex = rich.IsFormatted
                ? context.InternRichString(rich, stored.Markup, context.StyleTable.FontOfXf(xf))
                : context.InternString(text);
        }
        else
        {
            sstIndex = context.InternString(text);
        }

        int start = writer.BeginRecord(BiffRecordType.LabelSst);
        AppendCellHeader(writer, cell, xf);
        writer.AppendUInt32((uint)sstIndex);
        writer.EndRecord(start);
    }

    private void WriteFormulaCell(BiffWriter writer, CellRef cell, int xf, string formula, in CellValue value)
    {
        byte[]? rgce = _encoder.Encode(formula, cell);
        if (rgce is null)
        {
            context.Warn(new WorkbookWarning(
                "UnwritableFormula",
                $"The formula in '{sheet.Name}' has no encoding in the .xls format; its last value was written instead.",
                cell));
            WriteValueCell(writer, cell, xf, value);
            return;
        }

        int start = writer.BeginRecord(BiffRecordType.Formula);
        AppendCellHeader(writer, cell, xf);

        Span<byte> cached = stackalloc byte[8];
        BiffValues.WriteCachedValue(cached, value, context.Workbook.DateSystem);
        writer.Append(cached);

        writer.AppendUInt16(0x0002);            // Calculate on load.
        writer.AppendUInt32(0);                 // Reserved.
        writer.AppendUInt16(rgce.Length);
        writer.Append(rgce);
        writer.EndRecord(start);

        // A formula whose result is text carries the text in the record that follows it.
        if (value.Kind == CellValueKind.Text && value.GetText().Length > 0)
        {
            int stringStart = writer.BeginRecord(BiffRecordType.String);
            writer.AppendUnicodeString(value.GetText());
            writer.EndRecord(stringStart);
        }
    }

    private static void WriteCellHeader(BiffWriter writer, ushort type, CellRef cell, int xf, int bodySize)
    {
        Span<byte> body = stackalloc byte[6 + bodySize];
        body.Clear();
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, (ushort)(cell.Row - 1));
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[2..], (ushort)(cell.Column - 1));
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body[4..], (ushort)xf);
        writer.WriteRecord(type, body);
    }

    private static void AppendCellHeader(BiffWriter writer, CellRef cell, int xf)
    {
        writer.AppendUInt16(cell.Row - 1);
        writer.AppendUInt16(cell.Column - 1);
        writer.AppendUInt16(xf);
    }

    /// <summary>
    /// Writes the index that closes a block: how far back the first Row record is, then the
    /// distance from each row's cells to the next row's.
    /// </summary>
    private static void WriteDbCell(BiffWriter writer, int firstRowRecord, List<int> cellOffsets)
    {
        int start = writer.BeginRecord(BiffRecordType.DbCell);
        writer.AppendUInt32((uint)(writer.Length - 4 - firstRowRecord));
        for (int i = 0; i < cellOffsets.Count; i++)
        {
            int next = i + 1 < cellOffsets.Count ? cellOffsets[i + 1] : start - 4;
            writer.AppendUInt16(Math.Clamp(next - cellOffsets[i], 0, ushort.MaxValue));
        }

        writer.EndRecord(start);
    }

    private void WriteWindow(BiffWriter writer)
    {
        int flags = 0x02B0;                       // Defaults: headings, zeros, outline symbols.
        if (sheet.ShowGridLines)
            flags |= 0x0002;
        if (sheetIndex == 0)
            flags |= 0x0400 | 0x0200;             // Selected and active.
        if (sheet.FrozenRows > 0 || sheet.FrozenColumns > 0)
            flags |= 0x0008 | 0x0100;             // Frozen, and frozen without splitting.

        int start = writer.BeginRecord(BiffRecordType.Window2);
        writer.AppendUInt16(flags);
        writer.AppendUInt16(0);                   // Top row.
        writer.AppendUInt16(0);                   // Left column.
        writer.AppendUInt32(0x00000040);          // Automatic grid colour.
        writer.AppendUInt16(0);                   // Page break preview magnification.
        writer.AppendUInt16(0);                   // Normal view magnification.
        writer.AppendUInt32(0);                   // Reserved.
        writer.EndRecord(start);

        if (sheet.ZoomScale > 0 && sheet.ZoomScale != 100)
        {
            int sclStart = writer.BeginRecord(BiffRecordType.Scl);
            writer.AppendUInt16(sheet.ZoomScale);
            writer.AppendUInt16(100);
            writer.EndRecord(sclStart);
        }

        if (sheet.FrozenRows == 0 && sheet.FrozenColumns == 0)
            return;

        int paneStart = writer.BeginRecord(BiffRecordType.Pane);
        writer.AppendUInt16(sheet.FrozenColumns);
        writer.AppendUInt16(sheet.FrozenRows);
        writer.AppendUInt16(sheet.FrozenRows);
        writer.AppendUInt16(sheet.FrozenColumns);
        writer.AppendUInt16(2);                   // Active pane: bottom-right.
        writer.EndRecord(paneStart);
    }

    /// <summary>
    /// Writes the merged ranges. One record holds 1027 of them at most, so a heavily merged
    /// sheet needs several.
    /// </summary>
    private void WriteMerges(BiffWriter writer)
    {
        const int PerRecord = 1027;
        IReadOnlyList<RangeRef> merges = sheet.MergedRanges;
        for (int offset = 0; offset < merges.Count; offset += PerRecord)
        {
            int count = Math.Min(PerRecord, merges.Count - offset);
            int start = writer.BeginRecord(BiffRecordType.MergeCells);
            writer.AppendUInt16(count);
            for (int i = 0; i < count; i++)
            {
                RangeRef range = merges[offset + i];
                writer.AppendUInt16(Math.Min(range.Start.Row, MaxRow) - 1);
                writer.AppendUInt16(Math.Min(range.End.Row, MaxRow) - 1);
                writer.AppendUInt16(Math.Min(range.Start.Column, MaxColumn) - 1);
                writer.AppendUInt16(Math.Min(range.End.Column, MaxColumn) - 1);
            }

            writer.EndRecord(start);
        }
    }

    private void WriteHyperlinks(BiffWriter writer)
    {
        foreach (Hyperlink hyperlink in sheet.Hyperlinks)
            HyperlinkWriter.Write(writer, hyperlink);
    }

    private void WriteProtection(BiffWriter writer)
    {
        if (!sheet.Protection.IsEnabled)
            return;

        Span<byte> body = stackalloc byte[2];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(body, 1);
        writer.WriteRecord(BiffRecordType.Protect, body);
    }

    /// <summary>Writes the sheet's conditional formatting rules.</summary>
    private void WriteConditionalFormats(BiffWriter writer) =>
        ConditionalFormatWriter.Write(writer, sheet, _encoder, context);

    private void WriteValidations(BiffWriter writer) =>
        DataValidationWriter.Write(writer, sheet, _encoder, context.Warn);

    /// <summary>
    /// Writes the sheet's drawing layer — notes and pictures — claiming a drawing for them the
    /// first time either appears.
    /// </summary>
    private void WriteNotes(BiffWriter writer)
    {
        if (sheet.Comments.Count == 0 && sheet.Pictures.Count == 0)
            return;

        int drawingId = context.NextDrawingId();
        int shapes = EscherDrawingWriter.Write(writer, sheet, drawingId, context);
        if (shapes > 0)
            context.RegisterDrawing(drawingId, shapes);
    }
}
