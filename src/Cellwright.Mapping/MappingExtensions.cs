using System.Runtime.CompilerServices;

namespace Cellwright.Mapping;

/// <summary>
/// Typed reading and writing over the streaming layer and the DOM, driven entirely by
/// generated <see cref="IRowSchema{TRow}"/> implementations.
/// </summary>
public static class MappingExtensions
{
    /// <summary>
    /// Streams a sheet as typed rows. With <c>HasHeader</c> schemas the first row binds
    /// columns by header text (case-insensitive); otherwise columns map positionally.
    /// </summary>
    public static async IAsyncEnumerable<TRow> ReadRowsAsync<TRow>(
        this XlsxReader reader, string sheetName,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where TRow : IRowSchema<TRow>
    {
        ArgumentNullException.ThrowIfNull(reader);
        var columnMap = new int[TRow.ColumnCount];
        bool bound = !TRow.HasHeader;
        if (bound)
        {
            for (int i = 0; i < columnMap.Length; i++)
                columnMap[i] = i;
        }

        await foreach (RowData row in reader.ReadRowsAsync(sheetName, cancellationToken).ConfigureAwait(false))
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

    /// <summary>Writes typed rows to a streaming sheet, emitting the header row when the schema has one.</summary>
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
            for (int i = 0; i < TRow.Headers.Length; i++)
                cells[i] = TRow.Headers[i];
            sheet.WriteRow(cells);
        }

        foreach (TRow row in rows)
        {
            TRow.Write(row, cells);
            sheet.WriteRow(cells, styles);
            await sheet.FlushIfNeededAsync(cancellationToken).ConfigureAwait(false);
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
            for (int i = 0; i < cells.Length; i++)
            {
                Cell target = sheet.Cell(row, start.Column + i).SetValue(cells[i]);
                if (styles[i] != StyleId.Default)
                    target.SetStyle(styles[i]);
            }

            row++;
        }

        return sheet;
    }

    private static void BindHeader<TRow>(RowData headerRow, int[] columnMap) where TRow : IRowSchema<TRow>
    {
        ReadOnlySpan<string> expected = TRow.Headers;
        List<string>? missing = null;
        for (int i = 0; i < expected.Length; i++)
        {
            columnMap[i] = FindColumn(headerRow.Cells, expected[i]);
            if (columnMap[i] < 0)
                (missing ??= []).Add(expected[i]);
        }

        if (missing is not null)
        {
            throw new CellMappingException(
                $"Header row {headerRow.RowNumber} is missing expected column(s): {string.Join(", ", missing)}.");
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
