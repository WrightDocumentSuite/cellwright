namespace Cellwright;

/// <summary>
/// The print area and the rows or columns repeated on every page. Excel keeps all three in
/// sheet-scoped built-in names rather than in the worksheet part, so these read and write
/// <c>_xlnm.Print_Area</c> and <c>_xlnm.Print_Titles</c> for you.
/// </summary>
public static class PrintTitles
{
    private const string PrintAreaName = "_xlnm.Print_Area";
    private const string PrintTitlesName = "_xlnm.Print_Titles";

    extension(Worksheet sheet)
    {
        /// <summary>Range printed from this sheet, or <see langword="null"/> to print everything.</summary>
        public RangeRef? PrintArea
        {
            get
            {
                string? refersTo = sheet.Workbook.Names.Find(PrintAreaName, sheet)?.RefersTo;
                return TryStripSheet(refersTo, out ReadOnlySpan<char> range) &&
                       RangeRef.TryParse(range, null, out RangeRef parsed)
                    ? parsed
                    : null;
            }

            set
            {
                if (value is not { } range)
                {
                    sheet.Workbook.Names.Remove(PrintAreaName, sheet);
                    return;
                }

                sheet.Workbook.Names.Add(PrintAreaName, Qualify(sheet, range), sheet);
            }
        }

        /// <summary>Sets the print area from an A1 range such as <c>A1:H40</c>.</summary>
        public void SetPrintArea(string a1) => sheet.PrintArea = RangeRef.Parse(a1);

        /// <summary>Repeats a band of rows at the top of every printed page.</summary>
        public void SetRepeatRows(int first, int last)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(first, 1);
            ArgumentOutOfRangeException.ThrowIfLessThan(last, first);
            SetTitles(sheet, RowTitles(sheet, first, last), keepColumns: true);
        }

        /// <summary>Repeats a band of columns on the left of every printed page.</summary>
        public void SetRepeatColumns(int first, int last)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(first, 1);
            ArgumentOutOfRangeException.ThrowIfLessThan(last, first);
            SetTitles(sheet, ColumnTitles(sheet, first, last), keepColumns: false);
        }

        /// <summary>Stops repeating any rows or columns.</summary>
        public void ClearRepeatTitles() => sheet.Workbook.Names.Remove(PrintTitlesName, sheet);

        /// <summary>Rows repeated at the top of every page, as 1-based bounds.</summary>
        public (int First, int Last)? RepeatRows => FindTitles(sheet, rows: true);

        /// <summary>Columns repeated on the left of every page, as 1-based bounds.</summary>
        public (int First, int Last)? RepeatColumns => FindTitles(sheet, rows: false);
    }

    /// <summary>
    /// Writes one half of <c>Print_Titles</c> while keeping the other. The name holds the two
    /// halves as a comma-separated pair, columns first, and either may be missing.
    /// </summary>
    private static void SetTitles(Worksheet sheet, string replacement, bool keepColumns)
    {
        (int First, int Last)? kept = FindTitles(sheet, rows: !keepColumns);
        string? other = kept is not { } bounds
            ? null
            : keepColumns
                ? ColumnTitles(sheet, bounds.First, bounds.Last)
                : RowTitles(sheet, bounds.First, bounds.Last);

        string value = other is null
            ? replacement
            : keepColumns ? $"{other},{replacement}" : $"{replacement},{other}";
        sheet.Workbook.Names.Add(PrintTitlesName, value, sheet);
    }

    private static (int First, int Last)? FindTitles(Worksheet sheet, bool rows)
    {
        if (sheet.Workbook.Names.Find(PrintTitlesName, sheet)?.RefersTo is not { } refersTo)
            return null;

        foreach (Range segment in refersTo.AsSpan().Split(','))
        {
            if (!TryStripSheet(refersTo.AsSpan()[segment], out ReadOnlySpan<char> part))
                continue;

            int separator = part.IndexOf(':');
            if (separator < 0)
                continue;

            ReadOnlySpan<char> low = part[..separator].TrimStart('$');
            ReadOnlySpan<char> high = part[(separator + 1)..].TrimStart('$');
            bool isRow = !low.IsEmpty && char.IsAsciiDigit(low[0]);
            if (isRow != rows)
                continue;

            if (rows && TryRow(low, out int firstRow) && TryRow(high, out int lastRow))
                return (firstRow, lastRow);
            if (!rows && TryColumn(low, out int firstColumn) && TryColumn(high, out int lastColumn))
                return (firstColumn, lastColumn);
        }

        return null;
    }

    private static bool TryColumn(ReadOnlySpan<char> text, out int column)
    {
        int position = 0;
        return ColumnName.TryParse(text, ref position, out column) && position == text.Length;
    }

    /// <summary>A row number out of a defined name; invariant, as everything in a file is.</summary>
    private static bool TryRow(ReadOnlySpan<char> text, out int row) => int.TryParse(
        text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out row);

    private static string RowTitles(Worksheet sheet, int first, int last) =>
        $"{FormulaText.QuoteSheetName(sheet.Name)}!${first}:${last}";

    private static string ColumnTitles(Worksheet sheet, int first, int last) =>
        $"{FormulaText.QuoteSheetName(sheet.Name)}!${ColumnName.Format(first)}:${ColumnName.Format(last)}";

    private static string Qualify(Worksheet sheet, RangeRef range) =>
        $"{FormulaText.QuoteSheetName(sheet.Name)}!${ColumnName.Format(range.Start.Column)}${range.Start.Row}" +
        $":${ColumnName.Format(range.End.Column)}${range.End.Row}";

    /// <summary>Drops the <c>Sheet!</c> prefix of a built-in name's target.</summary>
    private static bool TryStripSheet(ReadOnlySpan<char> refersTo, out ReadOnlySpan<char> rest)
    {
        rest = refersTo;
        if (refersTo.IsEmpty)
            return false;

        int separator = refersTo.LastIndexOf('!');
        if (separator >= 0)
            rest = refersTo[(separator + 1)..];
        return !rest.IsEmpty;
    }

    private static bool TryStripSheet(string? refersTo, out ReadOnlySpan<char> rest)
    {
        rest = default;
        return refersTo is not null && TryStripSheet(refersTo.AsSpan(), out rest);
    }
}
