namespace Cellwright.Pdf.Layout;

/// <summary>The clickable areas a sheet's hyperlinks put over its cells.</summary>
internal sealed partial class CellPlacer
{
    /// <summary>Puts a clickable area over every cell a hyperlink covers on this page.</summary>
    public void Links()
    {
        foreach (Hyperlink link in _sheet.Hyperlinks)
        {
            if (link.Target is { Length: > 0 } url)
            {
                Cover(link.Range, url, targetSheet: null, targetCell: null);
                continue;
            }

            if (link.Location is { Length: > 0 } location &&
                TryResolveInternalLocation(location, out Worksheet? targetSheet, out CellRef targetCell))
            {
                Cover(link.Range, url: null, targetSheet, targetCell);
            }
            else
            {
                _context.Diagnostics.Add(
                    PdfExportWarningKind.ContentSkipped,
                    "A hyperlink that points inside the workbook is not clickable in the PDF.",
                    "internal-link");
            }
        }
    }

    private bool TryResolveInternalLocation(string location, out Worksheet? targetSheet, out CellRef targetCell)
    {
        targetSheet = null;
        targetCell = default;

        ReadOnlySpan<char> span = location.AsSpan().Trim();
        if (span.StartsWith('#'))
            span = span[1..].Trim();
        if (span.IsEmpty)
            return false;

        if (!TryParseSheetAndRange(span, _sheet, out targetSheet, out RangeRef range))
        {
            DefinedName? defined = _sheet.Workbook.Names.Resolve(span.ToString(), _sheet);
            if (defined is null || !TryParseSheetAndRange(defined.RefersTo.AsSpan().TrimStart('='), defined.Scope ?? _sheet, out targetSheet, out range))
                return false;
        }

        targetCell = range.Start;
        return targetSheet is not null;
    }

    private bool TryParseSheetAndRange(
        ReadOnlySpan<char> reference,
        Worksheet defaultSheet,
        out Worksheet? sheet,
        out RangeRef range)
    {
        sheet = defaultSheet;
        range = default;

        ReadOnlySpan<char> cellPart = reference;
        int bang = LastUnquotedBang(reference);
        if (bang >= 0)
        {
            ReadOnlySpan<char> rawSheet = reference[..bang].Trim();
            cellPart = reference[(bang + 1)..].Trim();
            string sheetName = UnquoteSheetName(rawSheet);
            sheet = _sheet.Workbook.Sheets.FirstOrDefault(
                s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
                return false;
        }

        if (RangeRef.TryParse(cellPart, null, out range))
            return true;

        if (CellRef.TryParse(cellPart, null, out CellRef single))
        {
            range = new RangeRef(single, single);
            return true;
        }

        return false;
    }

    private static int LastUnquotedBang(ReadOnlySpan<char> text)
    {
        bool inQuotes = false;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '\'')
            {
                if (inQuotes && i + 1 < text.Length && text[i + 1] == '\'')
                {
                    i++;
                    continue;
                }

                inQuotes = !inQuotes;
            }
            else if (text[i] == '!' && !inQuotes)
            {
                return i;
            }
        }

        return -1;
    }

    private static string UnquoteSheetName(ReadOnlySpan<char> raw)
    {
        if (raw.Length >= 2 && raw[0] == '\'' && raw[^1] == '\'')
            return raw[1..^1].ToString().Replace("''", "'", StringComparison.Ordinal);

        return raw.ToString();
    }

    /// <summary>Covers whatever part of a link's range this page shows.</summary>
    private void Cover(RangeRef range, string? url, Worksheet? targetSheet, CellRef? targetCell)
    {
        for (int r = 0; r < _rows.Count; r++)
        {
            int row = _rows.Lines[r];
            if (row < range.Start.Row || row > range.End.Row)
                continue;

            for (int c = 0; c < _columns.Count; c++)
            {
                int column = _columns.Lines[c];
                if (column < range.Start.Column || column > range.End.Column)
                    continue;

                _page.Items.Add(new LinkItem
                {
                    X = _columns.Near[c],
                    Y = _rows.Near[r],
                    Width = _columns.Size[c],
                    Height = _rows.Size[r],
                    Url = url,
                    TargetSheet = targetSheet,
                    TargetCell = targetCell,
                });
            }
        }
    }
}
