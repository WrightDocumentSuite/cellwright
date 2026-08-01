using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;
using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright;

public sealed partial class XlsxReader
{
    /// <summary>Enumerates the rows of a sheet located by name (case-insensitive).</summary>
    public IAsyncEnumerable<RowData> ReadRowsAsync(string sheetName, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(sheetName);
        foreach (SheetInfo sheet in _sheets)
        {
            if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return ReadRowsAsync(sheet.Index, cancellationToken);
        }

        throw new ArgumentException($"The workbook has no sheet named '{sheetName}'.", nameof(sheetName));
    }

    /// <summary>
    /// Enumerates the rows of a sheet by zero-based index. Rows arrive in file order;
    /// empty rows are skipped (watch <see cref="RowData.RowNumber"/>). The yielded
    /// <see cref="RowData"/> is a view over a reused buffer — copy it to keep it.
    /// </summary>
    public async IAsyncEnumerable<RowData> ReadRowsAsync(
        int sheetIndex, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfNegative(sheetIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(sheetIndex, _sheets.Count);

        await EnsureSharedStringsAsync(cancellationToken).ConfigureAwait(false);
        await EnsureStylesAsync(cancellationToken).ConfigureAwait(false);

        var buffer = new CellValue[64];
        int previousCount = 0;
        int lastRowNumber = 0;

        Stream stream = await _package.OpenPartAsync(_sheetParts[sheetIndex], cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            using var xml = XmlReader.Create(stream, XmlDefaults.AsyncReaderSettings);
            while (await xml.ReadAsync().ConfigureAwait(false))
            {
                if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "row")
                    continue;

                cancellationToken.ThrowIfCancellationRequested();
                int rowNumber = int.TryParse(xml.GetAttribute("r"), NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)
                    ? parsed
                    : lastRowNumber + 1;
                lastRowNumber = rowNumber;
                if (xml.IsEmptyElement)
                    continue;

                Array.Clear(buffer, 0, previousCount);
                (int count, buffer) = await ParseRowCellsAsync(xml, buffer, rowNumber).ConfigureAwait(false);
                previousCount = count;
                if (count > 0)
                    yield return new RowData(buffer, count, rowNumber);
            }
        }
    }

    /// <summary>Parses the cells of the current <c>row</c> element into the buffer, growing it as needed.</summary>
    private async ValueTask<(int Count, CellValue[] Buffer)> ParseRowCellsAsync(XmlReader xml, CellValue[] buffer, int rowNumber)
    {
        int maxColumn = 0;
        int lastColumn = 0;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "row")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;
            if (xml.LocalName != "c")
            {
                await xml.SkipAsync().ConfigureAwait(false);
                advance = false;
                continue;
            }

            int column = ResolveColumn(xml.GetAttribute("r"), lastColumn, rowNumber);
            lastColumn = column;
            if (column > SheetLimits.MaxColumn)
            {
                Warn("ColumnOutOfRange", $"A cell beyond column XFD was skipped in row {rowNumber}.");
                if (!xml.IsEmptyElement)
                {
                    await xml.SkipAsync().ConfigureAwait(false);
                    advance = false;
                }

                continue;
            }

            int styleIndex = int.TryParse(xml.GetAttribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out int style) ? style : 0;
            string? type = xml.GetAttribute("t");
            var cell = new CellRef(rowNumber, column);

            CellValue value = xml.IsEmptyElement
                ? CellValue.Blank
                : (await ParseCellContentAsync(xml, type, styleIndex, cell, captureFormula: false).ConfigureAwait(false)).Value;

            if (value.IsBlank)
                continue;


            if (column > buffer.Length)
                Array.Resize(ref buffer, Math.Max(buffer.Length * 2, column));
            buffer[column - 1] = value;
            maxColumn = Math.Max(maxColumn, column);
        }

        return (maxColumn, buffer);
    }

    private int ResolveColumn(string? cellRef, int lastColumn, int rowNumber)
    {
        if (cellRef is null)
            return lastColumn + 1;

        if (CellRef.TryParse(cellRef, null, out CellRef parsed))
            return parsed.Column;

        Warn("BadCellRef", $"Cell reference '{cellRef}' in row {rowNumber} is malformed; using the next column.");
        return lastColumn + 1;
    }

    /// <summary>Parses the children of a <c>c</c> element; the reader ends on its end tag.</summary>
    private async ValueTask<(CellValue Value, ParsedFormula Formula, int? SharedStringIndex)> ParseCellContentAsync(
        XmlReader xml, string? type, int styleIndex, CellRef cell, bool captureFormula)
    {
        CellValue value = CellValue.Blank;
        ParsedFormula formula = default;
        int? sharedStringIndex = null;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "c")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "v":
                    string text = await xml.ReadElementContentAsStringAsync().ConfigureAwait(false);
                    advance = false;
                    value = InterpretValue(type, text, styleIndex, cell);
                    if (type == "s" && int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int index) &&
                        (uint)index < (uint)_sharedStrings!.Length)
                    {
                        sharedStringIndex = index;
                    }
                    break;

                case "is":
                    value = CellValue.Text(await ReadStringItemAsync(xml, "is").ConfigureAwait(false));
                    break;

                case "f" when captureFormula:
                    bool empty = xml.IsEmptyElement;
                    formula = await ParseFormulaAsync(xml).ConfigureAwait(false);
                    advance = empty;
                    break;

                default: // f (when not captured), extLst and friends carry no cell value.
                    await xml.SkipAsync().ConfigureAwait(false);
                    advance = false;
                    break;
            }
        }

        return (value, formula, sharedStringIndex);
    }

    /// <summary>
    /// Reads a formula element together with the attributes that describe shared and array
    /// groups. Group expansion happens later (in the DOM loader), so nothing is lost here.
    /// </summary>
    private static async ValueTask<ParsedFormula> ParseFormulaAsync(XmlReader xml)
    {
        string? formulaType = xml.GetAttribute("t");
        string? rangeText = xml.GetAttribute("ref");
        int? sharedIndex = int.TryParse(
            xml.GetAttribute("si"), NumberStyles.None, CultureInfo.InvariantCulture, out int si)
            ? si
            : null;

        // A data table carries no text, only attributes naming its two input cells, so the
        // whole element is kept. Every other kind keeps just the attributes the model has no
        // room for — above all "ca", which marks a formula Excel must always recalculate.
        bool dataTable = formulaType == "dataTable";
        List<KeyValuePair<string, string>>? attributes = dataTable
            ? RawXml.ReadAttributes(xml)
            : RawXml.ReadAttributes(xml, "t", "ref", "si");
        FormulaMarkup? markup = attributes is null ? null : new FormulaMarkup(attributes, dataTable);

        string text = xml.IsEmptyElement ? "" : await xml.ReadElementContentAsStringAsync().ConfigureAwait(false);

        RangeRef? range = RangeRef.TryParse(rangeText, null, out RangeRef parsed) ? parsed : null;
        return new ParsedFormula(
            text.Length > 0 ? OoxmlText.Decode(text) : null, formulaType, sharedIndex, range, markup);
    }

    private CellValue InterpretValue(string? type, string text, int styleIndex, CellRef cell)
    {
        if (text.Length == 0)
            return CellValue.Blank;

        switch (type)
        {
            case null or "n":
                return InterpretNumber(text, styleIndex, cell);

            case "s":
                if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int index) &&
                    (uint)index < (uint)_sharedStrings!.Length)
                {
                    return CellValue.Text(_sharedStrings[index]);
                }

                Warn("BadSharedStringIndex", $"Shared string index '{text}' is out of range.", cell);
                return CellValue.Blank;

            case "b":
                if (text == "1")
                    return CellValue.Boolean(true);
                if (text == "0")
                    return CellValue.Boolean(false);
                if (bool.TryParse(text, out bool logical))
                    return CellValue.Boolean(logical);
                Warn("BadBoolean", $"Boolean cell value '{text}' is malformed.", cell);
                return CellValue.Blank;

            case "e":
                if (ExcelErrorText.TryParse(text, out ExcelError error))
                    return CellValue.Error(error);
                Warn("UnknownError", $"Unknown error literal '{text}'.", cell);
                return CellValue.Error(ExcelError.Value);

            case "str":
                return CellValue.Text(OoxmlText.Decode(text));

            case "d":
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime isoDate))
                    return CellValue.DateTime(isoDate);
                Warn("BadIsoDate", $"ISO date cell value '{text}' is malformed.", cell);
                return CellValue.Blank;

            default:
                Warn("UnknownCellType", $"Unknown cell type '{type}'; treating the value as text.", cell);
                return CellValue.Text(text);
        }
    }

    private CellValue InterpretNumber(string text, int styleIndex, CellRef cell)
    {
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            Warn("BadNumber", $"Numeric cell value '{text}' is malformed.", cell);
            return CellValue.Blank;
        }

        TemporalKind kind = (uint)styleIndex < (uint)_styleTemporalKinds!.Length
            ? _styleTemporalKinds[styleIndex]
            : TemporalKind.None;

        switch (kind)
        {
            case TemporalKind.Date when number >= 0:
                try
                {
                    return CellValue.DateTime(ExcelSerialDate.ToDateTime(number, DateSystem));
                }
                catch (ArgumentException)
                {
                    Warn("BadDateSerial", $"Serial value {number} does not map to a valid date.", cell);
                    return CellValue.Number(number);
                }

            case TemporalKind.Duration:
                return CellValue.TimeSpan(ExcelSerialDate.ToTimeSpan(number));

            default:
                return CellValue.Number(number);
        }
    }
}
