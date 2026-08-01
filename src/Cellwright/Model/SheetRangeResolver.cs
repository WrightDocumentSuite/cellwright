namespace Cellwright;

/// <summary>
/// Resolves a sheet-qualified range written as text — <c>Лист1!$B$2:$B$10</c>,
/// <c>'Q1 2026'!A1</c>, or a defined name standing for either — to the sheet and rectangle it
/// names.
/// </summary>
/// <remarks>
/// This is not a formula evaluator and takes care not to grow into one: the only things it
/// accepts are a qualified reference and a name that resolves to one, because those are the
/// only things the places that need it — chart series, print areas, pivot sources — are allowed
/// to hold. Anything else answers null rather than a guess.
/// </remarks>
internal static class SheetRangeResolver
{
    /// <summary>How many name-to-name hops are followed before giving up.</summary>
    private const int MaxNameDepth = 4;

    /// <summary>Resolves a reference to the sheet and rectangle it names, or null.</summary>
    /// <param name="workbook">The workbook the reference speaks of.</param>
    /// <param name="reference">The text, with or without a leading <c>=</c>.</param>
    public static (Worksheet Sheet, RangeRef Range)? Resolve(Workbook workbook, string reference)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        if (string.IsNullOrWhiteSpace(reference))
            return null;

        return Resolve(workbook, reference.AsSpan().Trim(), depth: 0);
    }

    private static (Worksheet, RangeRef)? Resolve(Workbook workbook, ReadOnlySpan<char> text, int depth)
    {
        if (text.Length > 0 && text[0] == '=')
            text = text[1..].Trim();
        if (text.IsEmpty)
            return null;

        int position = 0;
        string? sheetName = null;
        if (ReferenceScanner.TryScanSheetPrefix(text, ref position, out string parsed))
            sheetName = parsed;

        if (ReferenceScanner.TryScanPart(text, ref position, out RefPart start))
        {
            RefPart end = start;
            if (position < text.Length && text[position] == ':')
            {
                position++;
                if (!ReferenceScanner.TryScanPart(text, ref position, out end) || end.Kind != start.Kind)
                    return null;
            }

            // Only the whole text counts: "A1:B2 something" is not a reference plus noise, it
            // is not a reference.
            if (position != text.Length)
                return null;

            Worksheet? sheet = sheetName is null
                ? workbook.Sheets.Count > 0 ? workbook.Sheets[0] : null
                : workbook.Sheets.Find(sheetName);

            if (sheet is null)
                return null;

            return Rectangle(sheet, start, end);
        }

        // Not a reference: perhaps a defined name that stands for one.
        if (sheetName is not null || depth >= MaxNameDepth)
            return null;

        DefinedName? name = workbook.Names.Resolve(text.ToString(), scope: null);
        return name is null ? null : Resolve(workbook, name.RefersTo.AsSpan().Trim(), depth + 1);
    }

    /// <summary>The rectangle two parts span, whole rows and columns expanded to the grid.</summary>
    private static (Worksheet, RangeRef)? Rectangle(Worksheet sheet, in RefPart start, in RefPart end)
    {
        (int firstRow, int firstColumn, int lastRow, int lastColumn) = start.Kind switch
        {
            RefPartKind.Cell => (start.Row, start.Column, end.Row, end.Column),
            RefPartKind.Column => (1, start.Column, SheetLimits.MaxRow, end.Column),
            RefPartKind.Row => (start.Row, 1, end.Row, SheetLimits.MaxColumn),
            _ => (0, 0, 0, 0),
        };

        if (firstRow <= 0 || firstColumn <= 0 ||
            !SheetLimits.IsValidRow(firstRow) || !SheetLimits.IsValidRow(lastRow) ||
            !SheetLimits.IsValidColumn(firstColumn) || !SheetLimits.IsValidColumn(lastColumn))
        {
            return null;
        }

        // Written backwards is still the same rectangle.
        if (lastRow < firstRow)
            (firstRow, lastRow) = (lastRow, firstRow);
        if (lastColumn < firstColumn)
            (firstColumn, lastColumn) = (lastColumn, firstColumn);

        return (sheet, new RangeRef(new CellRef(firstRow, firstColumn), new CellRef(lastRow, lastColumn)));
    }
}
