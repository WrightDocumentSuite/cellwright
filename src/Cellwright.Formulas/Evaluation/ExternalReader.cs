using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Answers references into other workbooks from their cached values. A single cell arrives as
/// a scalar and anything larger as an array — there is no worksheet behind the cache for a
/// real reference to point at.
/// </summary>
internal static class ExternalReader
{
    public static FormulaValue Read(ExternalBook book, in FormulaRef reference, EvalContext context)
    {
        if (reference.ExternalSheetName is not { } sheetName)
            return FormulaValue.Error(ExcelError.Reference);

        int firstSheet = book.IndexOfSheet(sheetName);
        int lastSheet = reference.ExternalSpanEnd is { } spanEnd ? book.IndexOfSheet(spanEnd) : firstSheet;
        if (firstSheet < 0 || lastSheet < 0)
            return FormulaValue.Error(ExcelError.Reference);
        if (firstSheet > lastSheet)
            (firstSheet, lastSheet) = (lastSheet, firstSheet);

        if (!TryCorners(reference, book, firstSheet, lastSheet,
                out int firstRow, out int firstColumn, out int lastRow, out int lastColumn))
        {
            return FormulaValue.Error(ExcelError.Reference);
        }

        _ = context;
        if (firstSheet == lastSheet && firstRow == lastRow && firstColumn == lastColumn)
            return FormulaValue.Scalar(book.Sheets[firstSheet].Read(firstRow, firstColumn));

        int rows = lastRow - firstRow + 1;
        int columns = lastColumn - firstColumn + 1;
        int sheetCount = lastSheet - firstSheet + 1;
        long total = (long)rows * columns * sheetCount;
        if (total > 4_000_000)
            return FormulaValue.Error(ExcelError.Reference);

        // A span stacks its sheets on top of each other; aggregates read the values and never
        // ask which sheet a row came from.
        var values = new CellValue[total];
        int at = 0;
        for (int sheet = firstSheet; sheet <= lastSheet; sheet++)
        {
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int column = firstColumn; column <= lastColumn; column++)
                    values[at++] = book.Sheets[sheet].Read(row, column);
            }
        }

        return FormulaValue.Array(new ArrayValue(rows * sheetCount, columns, values));
    }

    /// <summary>
    /// Concrete corners of the reference, whole columns and rows clamped to the extent the
    /// caches actually hold.
    /// </summary>
    private static bool TryCorners(
        in FormulaRef reference, ExternalBook book, int firstSheet, int lastSheet,
        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
    {
        int cachedRows = 1;
        int cachedColumns = 1;
        for (int sheet = firstSheet; sheet <= lastSheet; sheet++)
        {
            cachedRows = Math.Max(cachedRows, book.Sheets[sheet].LastRow);
            cachedColumns = Math.Max(cachedColumns, book.Sheets[sheet].LastColumn);
        }

        (firstRow, firstColumn) = Corner(reference.Start, cachedRows, cachedColumns, start: true);
        (lastRow, lastColumn) = reference.IsRange
            ? Corner(reference.End, cachedRows, cachedColumns, start: false)
            : (firstRow, firstColumn);

        if (firstRow > lastRow)
            (firstRow, lastRow) = (lastRow, firstRow);
        if (firstColumn > lastColumn)
            (firstColumn, lastColumn) = (lastColumn, firstColumn);

        return firstRow >= 1 && firstColumn >= 1 &&
            SheetLimits.IsValidRow(lastRow) && SheetLimits.IsValidColumn(lastColumn);
    }

    private static (int Row, int Column) Corner(RefPart part, int cachedRows, int cachedColumns, bool start) =>
        part.Kind switch
        {
            RefPartKind.Cell => (part.Row, part.Column),
            RefPartKind.Column => (start ? 1 : cachedRows, part.Column),
            RefPartKind.Row => (part.Row, start ? 1 : cachedColumns),
            _ => (0, 0),
        };
}
