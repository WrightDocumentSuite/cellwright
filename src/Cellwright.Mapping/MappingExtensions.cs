using System.Runtime.CompilerServices;

namespace Cellwright.Mapping;

/// <summary>
/// Typed reading and writing over the streaming layer (XLSX and CSV/TSV) and the DOM,
/// driven entirely by generated <see cref="IRowSchema{TRow}"/> implementations.
/// </summary>
public static class MappingExtensions
{
    /// <summary>
    /// Streams a sheet as typed rows. With <c>HasHeader</c> schemas the first row binds
    /// columns by header text (case-insensitive); otherwise columns map positionally.
    /// </summary>
    public static IAsyncEnumerable<TRow> ReadRowsAsync<TRow>(
        this XlsxReader reader, string sheetName,
        CancellationToken cancellationToken = default)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(reader);
        return ReadRowsCoreAsync<TRow>(reader.ReadRowsAsync(sheetName, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Streams typed rows from an asynchronous <see cref="RowData"/> sequence such as
    /// <see cref="CsvReader.ReadRowsAsync(Stream, CsvOptions?, CancellationToken)"/>.
    /// </summary>
    public static IAsyncEnumerable<TRow> ReadRowsAsync<TRow>(
        this IAsyncEnumerable<RowData> rows,
        CancellationToken cancellationToken = default)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(rows);
        return ReadRowsCoreAsync<TRow>(rows, cancellationToken);
    }

    /// <summary>
    /// Reads typed rows from a synchronous <see cref="RowData"/> sequence such as
    /// <see cref="CsvReader.ReadRows(TextReader, CsvOptions?)"/>.
    /// </summary>
    public static IEnumerable<TRow> ReadRows<TRow>(this IEnumerable<RowData> rows)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(rows);
        return ReadRowsSyncCore<TRow>(rows);
    }

    /// <summary>Writes typed rows to a streaming XLSX sheet, emitting the header row when the schema has one.</summary>
    public static async ValueTask WriteRowsAsync<TRow>(
        this XlsxSheetWriter sheet, IEnumerable<TRow> rows, CancellationToken cancellationToken = default)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(rows);

        var cells = new CellValue[TRow.ColumnCount];
        var styles = new StyleId[TRow.ColumnCount];
        TRow.RegisterStyles(sheet.Writer.RegisterStyle, styles);

        if (TRow.HasHeader)
        {
            FillHeaders<TRow>(cells);
            sheet.WriteRow(cells);
        }

        foreach (TRow row in rows)
        {
            TRow.Write(row, cells);
            sheet.WriteRow(cells, styles);
            await sheet.FlushIfNeededAsync(cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Writes typed rows to a <see cref="CsvWriter"/> asynchronously, emitting the header row when the schema has one.</summary>
    public static async ValueTask WriteRowsAsync<TRow>(
        this CsvWriter writer, IEnumerable<TRow> rows, CancellationToken cancellationToken = default)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(writer);
        ArgumentNullException.ThrowIfNull(rows);

        var cells = new CellValue[TRow.ColumnCount];
        if (TRow.HasHeader)
        {
            FillHeaders<TRow>(cells);
            await writer.WriteRowAsync(cells, cancellationToken).ConfigureAwait(false);
        }

        foreach (TRow row in rows)
        {
            TRow.Write(row, cells);
            await writer.WriteRowAsync(cells, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Writes typed rows to a <see cref="CsvWriter"/> synchronously, emitting the header row when the schema has one.</summary>
    public static void WriteRows<TRow>(this CsvWriter writer, IEnumerable<TRow> rows)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(writer);
        ArgumentNullException.ThrowIfNull(rows);

        var cells = new CellValue[TRow.ColumnCount];
        if (TRow.HasHeader)
        {
            FillHeaders<TRow>(cells);
            writer.WriteRow(cells);
        }

        foreach (TRow row in rows)
        {
            TRow.Write(row, cells);
            writer.WriteRow(cells);
        }
    }

    /// <summary>Inserts typed rows into a DOM worksheet starting at <paramref name="startCell"/>.</summary>
    public static Worksheet InsertTable<TRow>(this Worksheet sheet, IEnumerable<TRow> rows, string startCell = "A1")
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(rows);
        CellRef start = CellRef.Parse(startCell);

        var cells = new CellValue[TRow.ColumnCount];
        var styles = new StyleId[TRow.ColumnCount];
        TRow.RegisterStyles(sheet.Workbook.RegisterStyle, styles);

        int row = start.Row;
        if (TRow.HasHeader)
        {
            for (int i = 0; i < TRow.Headers.Length; i++)
                sheet.Cell(row, start.Column + i).SetValue(TRow.Headers[i]);
            row++;
        }

        foreach (TRow item in rows)
        {
            TRow.Write(item, cells);
            WriteDomRow(sheet, row++, start.Column, cells, styles);
        }

        return sheet;
    }

    /// <summary>Reads typed rows from a DOM worksheet starting at <paramref name="startCell"/> (or the top-left of <see cref="Worksheet.UsedRange"/>).</summary>
    public static IEnumerable<TRow> ReadTable<TRow>(this Worksheet sheet, string? startCell = null)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return ReadTableCore<TRow>(sheet, startCell);
    }

    private static async IAsyncEnumerable<TRow> ReadRowsCoreAsync<TRow>(
        IAsyncEnumerable<RowData> rows,
        [EnumeratorCancellation] CancellationToken cancellationToken)
        where TRow : IRowSchema<TRow>
    {
        int[] columnMap = CreateInitialColumnMap<TRow>(out bool bound);
        await foreach (RowData row in rows.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            if (!bound)
            {
                BindHeader<TRow>(row, columnMap);
                bound = true;
                continue;
            }

            yield return TRow.Read(row.Cells, columnMap);
        }
    }

    private static IEnumerable<TRow> ReadRowsSyncCore<TRow>(IEnumerable<RowData> rows)
        where TRow : IRowSchema<TRow>
    {
        int[] columnMap = CreateInitialColumnMap<TRow>(out bool bound);
        foreach (RowData row in rows)
        {
            if (!bound)
            {
                BindHeader<TRow>(row, columnMap);
                bound = true;
                continue;
            }

            yield return TRow.Read(row.Cells, columnMap);
        }
    }

    private static IEnumerable<TRow> ReadTableCore<TRow>(Worksheet sheet, string? startCell)
        where TRow : IRowSchema<TRow>
    {
        if (sheet.UsedRange is not { } used)
            yield break;

        CellRef start = startCell is null ? used.Start : CellRef.Parse(startCell);
        int width = Math.Max(TRow.ColumnCount, used.End.Column - start.Column + 1);
        var buffer = new CellValue[width];
        int[] columnMap = CreateInitialColumnMap<TRow>(out bool bound);

        for (int r = start.Row; r <= used.End.Row; r++)
        {
            int count = ReadDomRow(sheet, r, start.Column, width, buffer);
            if (count == 0)
                continue;

            ReadOnlySpan<CellValue> rowCells = buffer.AsSpan(0, count);
            if (!bound)
            {
                BindHeader<TRow>(rowCells, r, columnMap);
                bound = true;
                continue;
            }

            yield return TRow.Read(rowCells, columnMap);
        }
    }

    private static int ReadDomRow(Worksheet sheet, int row, int startCol, int width, CellValue[] buffer)
    {
        int lastNonBlank = 0;
        for (int i = 0; i < width; i++)
        {
            CellValue val = sheet.Cell(row, startCol + i).Value;
            buffer[i] = val;
            if (!val.IsBlank)
                lastNonBlank = i + 1;
        }

        return lastNonBlank;
    }

    private static void WriteDomRow(
        Worksheet sheet, int row, int startCol, ReadOnlySpan<CellValue> cells, ReadOnlySpan<StyleId> styles)
    {
        for (int i = 0; i < cells.Length; i++)
        {
            Cell target = sheet.Cell(row, startCol + i).SetValue(cells[i]);
            if (styles[i] != StyleId.Default)
                target.SetStyle(styles[i]);
        }
    }

    private static int[] CreateInitialColumnMap<TRow>(out bool bound) where TRow : IRowSchema<TRow>
    {
        var columnMap = new int[TRow.ColumnCount];
        bound = !TRow.HasHeader;
        if (bound)
        {
            for (int i = 0; i < columnMap.Length; i++)
                columnMap[i] = i;
        }

        return columnMap;
    }

    private static void FillHeaders<TRow>(Span<CellValue> cells) where TRow : IRowSchema<TRow>
    {
        for (int i = 0; i < TRow.Headers.Length; i++)
            cells[i] = TRow.Headers[i];
    }

    private static void BindHeader<TRow>(RowData headerRow, int[] columnMap) where TRow : IRowSchema<TRow> =>
        BindHeader<TRow>(headerRow.Cells, headerRow.RowNumber, columnMap);

    private static void BindHeader<TRow>(ReadOnlySpan<CellValue> headerCells, int rowNumber, int[] columnMap)
        where TRow : IRowSchema<TRow>
    {
        ReadOnlySpan<string> expected = TRow.Headers;
        List<string>? missing = null;
        for (int i = 0; i < expected.Length; i++)
        {
            columnMap[i] = FindColumn(headerCells, expected[i]);
            if (columnMap[i] < 0)
                (missing ??= []).Add(expected[i]);
        }

        if (missing is not null)
        {
            throw new CellMappingException(
                $"Header row {rowNumber} is missing expected column(s): {string.Join(", ", missing)}.");
        }
    }

    private static int FindColumn(ReadOnlySpan<CellValue> headerCells, string header)
    {
        for (int i = 0; i < headerCells.Length; i++)
        {
            if (headerCells[i].IsText &&
                string.Equals(headerCells[i].GetText().Trim(), header, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }
}
