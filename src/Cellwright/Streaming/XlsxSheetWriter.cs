using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright;

/// <summary>
/// Forward-only writer of a single worksheet. Rows are serialized immediately into the
/// package part; call <see cref="FlushIfNeededAsync"/> inside tight loops (or use
/// <see cref="WriteRowAsync"/>) to bound memory. Dispose to finalize the sheet.
/// </summary>
public sealed class XlsxSheetWriter : IAsyncDisposable
{
    private readonly XlsxWriter _owner;
    private readonly Utf8XmlWriter _xml;
    private readonly SheetOptions _options;
    private readonly List<RangeRef> _merges = [];
    private readonly IReadOnlyDictionary<int, RowProperties>? _rows;
    private readonly IReadOnlyDictionary<CellRef, RangeRef>? _arrayFormulaRanges;
    private readonly IReadOnlyDictionary<CellRef, FormulaMarkup>? _formulaMarkup;
    private readonly IReadOnlyDictionary<CellRef, List<KeyValuePair<string, string>>>? _cellMarkup;
    private readonly IReadOnlyDictionary<CellRef, RichStringMarkup>? _richStrings;
    private int _nextRow = 1;
    private bool _disposed;

    internal XlsxSheetWriter(XlsxWriter owner, Utf8XmlWriter xml, SheetOptions options)
    {
        _owner = owner;
        _xml = xml;
        _options = options;
        _rows = options.Rows;
        _arrayFormulaRanges = options.ArrayFormulaRanges;
        _formulaMarkup = options.FormulaMarkup;
        _cellMarkup = options.CellMarkup;
        _richStrings = options.RichStrings;
        WriteProlog(options);
    }

    /// <summary>The writer that owns this sheet (for registering styles mid-stream).</summary>
    public XlsxWriter Writer => _owner;

    /// <summary>The 1-based row number the next <c>WriteRow</c> call will produce.</summary>
    public int CurrentRow => _nextRow;

    /// <summary>
    /// Writes the next row. Blank values produce gaps; an entirely blank row advances the
    /// row counter without emitting an element. Values convert implicitly, so
    /// <c>sheet.WriteRow(1, "text", DateTime.Now)</c> works as-is.
    /// </summary>
    public void WriteRow(params scoped ReadOnlySpan<CellValue> cells) => WriteRow(cells, []);

    /// <summary>
    /// Writes the next row applying one style to every cell.
    /// Register styles via <see cref="XlsxWriter.RegisterStyle"/>.
    /// </summary>
    public void WriteRow(scoped ReadOnlySpan<CellValue> cells, StyleId style)
    {
        if (style == StyleId.Default)
        {
            WriteRow(cells, []);
            return;
        }

        Span<StyleId> styles = cells.Length <= 256 ? stackalloc StyleId[cells.Length] : new StyleId[cells.Length];
        styles.Fill(style);
        WriteRow(cells, styles);
    }

    /// <summary>
    /// Writes the next row with per-cell styles supplied in a parallel span. A shorter
    /// (or empty) span pads with the default style; blank cells with a style are written
    /// as empty styled cells so fills and borders still render.
    /// </summary>
    public void WriteRow(scoped ReadOnlySpan<CellValue> cells, scoped ReadOnlySpan<StyleId> styles)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_nextRow > SheetLimits.MaxRow)
            throw new InvalidOperationException($"The worksheet grid ends at row {SheetLimits.MaxRow}.");
        if (cells.Length > SheetLimits.MaxColumn)
            throw new ArgumentException($"A row cannot have more than {SheetLimits.MaxColumn} cells.", nameof(cells));
        if (styles.Length > cells.Length)
            throw new ArgumentException("The styles span cannot be longer than the cells span.", nameof(styles));

        int row = _nextRow++;
        if (AllBlank(cells) && styles.IsEmpty)
            return;

        OpenRowElement(row);
        for (int i = 0; i < cells.Length; i++)
        {
            int styleIndex = i < styles.Length ? styles[i].Index : 0;
            WriteCell(row, i + 1, in cells[i], styleIndex, formula: null);
        }

        _xml.WriteRaw("</row>"u8);
    }

    /// <summary>
    /// DOM save path: writes a row at an explicit 1-based position with raw style-pool
    /// indexes and formula texts. The position must not go backwards.
    /// </summary>
    internal void WriteRowAt(
        int rowNumber,
        scoped ReadOnlySpan<CellValue> cells,
        scoped ReadOnlySpan<int> styleIndexes,
        scoped ReadOnlySpan<string?> formulas)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentOutOfRangeException.ThrowIfLessThan(rowNumber, _nextRow);
        _nextRow = rowNumber + 1;

        OpenRowElement(rowNumber);
        for (int i = 0; i < cells.Length; i++)
        {
            WriteCell(
                rowNumber,
                i + 1,
                in cells[i],
                i < styleIndexes.Length ? styleIndexes[i] : 0,
                i < formulas.Length ? formulas[i] : null);
        }

        _xml.WriteRaw("</row>"u8);
    }

    private void OpenRowElement(int row)
    {
        _xml.WriteRaw("<row r=\""u8);
        _xml.WriteInt32(row);
        if (_rows is not null && _rows.TryGetValue(row, out RowProperties properties))
        {
            if (properties.Height is { } height)
            {
                _xml.WriteRaw("\" ht=\""u8);
                _xml.WriteDouble(height);
                _xml.WriteRaw("\" customHeight=\"1"u8);
            }

            if (properties.Hidden)
                _xml.WriteRaw("\" hidden=\"1"u8);
            if (properties.OutlineLevel > 0)
            {
                _xml.WriteRaw("\" outlineLevel=\""u8);
                _xml.WriteInt32(properties.OutlineLevel);
            }

            if (properties.Collapsed)
                _xml.WriteRaw("\" collapsed=\"1"u8);
            if (properties.Style != StyleId.Default)
            {
                _xml.WriteRaw("\" s=\""u8);
                _xml.WriteInt32(properties.Style.Index);
                _xml.WriteRaw("\" customFormat=\"1"u8);
            }

            if (properties.ThickTop)
                _xml.WriteRaw("\" thickTop=\"1"u8);
            if (properties.ThickBottom)
                _xml.WriteRaw("\" thickBot=\"1"u8);
        }

        _xml.WriteRaw("\">"u8);
    }

    /// <summary>Writes a row and flushes the buffer when it crossed the threshold.</summary>
    public ValueTask WriteRowAsync(ReadOnlyMemory<CellValue> cells, CancellationToken cancellationToken = default)
    {
        WriteRow(cells.Span);
        return _xml.FlushIfNeededAsync(cancellationToken);
    }

    /// <summary>Leaves the given number of rows empty.</summary>
    public void SkipRows(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        ObjectDisposedException.ThrowIf(_disposed, this);
        _nextRow += count;
    }

    /// <summary>Registers a merged range; merges are emitted when the sheet completes.</summary>
    public void Merge(RangeRef range)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _merges.Add(range);
    }

    /// <summary>Flushes buffered XML when the internal threshold was crossed. Cheap to call per row.</summary>
    public ValueTask FlushIfNeededAsync(CancellationToken cancellationToken = default) =>
        _xml.FlushIfNeededAsync(cancellationToken);

    /// <summary>Flushes all buffered XML to the package.</summary>
    public ValueTask FlushAsync(CancellationToken cancellationToken = default) =>
        _xml.FlushAsync(cancellationToken);

    /// <summary>
    /// Completes the worksheet. Everything after the cell data has to appear in the order
    /// CT_Worksheet prescribes, with the content we preserved slotted between the elements we
    /// generate — Excel and the schema validator both reject a sheet built any other way.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;
        _disposed = true;

        WorksheetMarkup? markup = _options.Markup;
        _xml.WriteRaw("</sheetData>"u8);
        RawXml.Write(_xml, markup?.BeforeAutoFilter);
        if (_options.Protection is { } protection)
            ProtectionPart.WriteSheet(_xml, protection);
        RawXml.Write(_xml, markup?.AfterProtection);
        WriteAutoFilter(markup);
        RawXml.Write(_xml, markup?.BeforeMerges);
        WriteMergeCells();
        RawXml.Write(_xml, markup?.AfterMerges);
        WriteConditionalFormats(markup);
        WriteValidations(markup);
        WriteHyperlinks();
        WritePageSetup(markup);
        RawXml.Write(_xml, markup?.PrintSetup);
        WriteDrawing(markup);
        WriteLegacyDrawing(markup);
        RawXml.Write(_xml, markup?.AfterLegacyDrawing);
        WriteTableParts(markup);
        RawXml.Write(_xml, markup?.Extensions);
        _xml.WriteRaw("</worksheet>"u8);
        await _xml.DisposeAsync().ConfigureAwait(false);
        _owner.OnSheetCompleted(this);
    }

    private void WriteProlog(SheetOptions options)
    {
        _xml.WriteDeclaration();
        _xml.WriteRaw("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""u8);
        _xml.WriteRaw(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""u8);
        RawXml.WriteAttributes(_xml, options.Markup?.RootAttributes);
        _xml.WriteRaw(">"u8);
        WriteSheetProperties(options);
        WriteSheetView(options);
        RawXml.Write(_xml, options.Markup?.SheetFormatProperties);
        WriteColumns(options);
        _xml.WriteRaw("<sheetData>"u8);
    }

    private void WriteSheetProperties(SheetOptions options)
    {
        WorksheetMarkup? markup = options.Markup;
        bool hasTabColor = !options.TabColor.IsAuto;
        bool ownsPageSetup = options.PageSetup is { IsModified: true };
        bool hasPageSetupProperties = ownsPageSetup
            ? options.PageSetup!.FitToPage || !options.PageSetup.AutoPageBreaks
            : !string.IsNullOrEmpty(markup?.PageSetupProperties);
        bool hasChildren = hasTabColor || hasPageSetupProperties || !string.IsNullOrEmpty(markup?.SheetPropertyChildren);
        if (!hasChildren && markup?.SheetProperties is null)
            return;

        _xml.WriteRaw("<sheetPr"u8);
        RawXml.WriteAttributes(_xml, markup?.SheetProperties);
        if (!hasChildren)
        {
            _xml.WriteRaw("/>"u8);
            return;
        }

        _xml.WriteRaw(">"u8);
        if (hasTabColor)
        {
            _xml.WriteRaw("<tabColor rgb=\""u8);
            _xml.WriteAttributeText(options.TabColor.ToHex());
            _xml.WriteRaw("\"/>"u8);
        }

        RawXml.Write(_xml, markup?.SheetPropertyChildren);
        if (ownsPageSetup)
            PageSetupWriter.WritePageSetupProperties(_xml, options.PageSetup!);
        else
            RawXml.Write(_xml, markup?.PageSetupProperties);

        _xml.WriteRaw("</sheetPr>"u8);
    }

    /// <summary>
    /// Writes the single sheet view we model: frozen panes, grid lines and zoom, with the
    /// selection and tab state of the source file replayed after the pane.
    /// </summary>
    private void WriteSheetView(SheetOptions options)
    {
        WorksheetMarkup? markup = options.Markup;
        int rows = options.FrozenRows;
        int columns = options.FrozenColumns;
        bool hasPane = rows > 0 || columns > 0;
        bool hasChildren = hasPane || !string.IsNullOrEmpty(markup?.SheetViewChildren);
        if (!hasChildren && options.ShowGridLines && options.ZoomScale == 100 && markup?.SheetViewAttributes is null)
            return;

        _xml.WriteRaw("<sheetViews><sheetView"u8);
        if (!options.ShowGridLines)
            _xml.WriteRaw(" showGridLines=\"0\""u8);
        if (options.ZoomScale != 100)
        {
            _xml.WriteRaw(" zoomScale=\""u8);
            _xml.WriteInt32(options.ZoomScale);
            _xml.WriteRaw("\""u8);
        }

        RawXml.WriteAttributes(_xml, markup?.SheetViewAttributes);
        if (!HasAttribute(markup?.SheetViewAttributes, "workbookViewId"))
            _xml.WriteRaw(" workbookViewId=\"0\""u8);

        if (!hasChildren)
        {
            _xml.WriteRaw("/></sheetViews>"u8);
            return;
        }

        _xml.WriteRaw(">"u8);
        if (hasPane)
            WriteFrozenPane(rows, columns);
        RawXml.Write(_xml, markup?.SheetViewChildren);
        _xml.WriteRaw("</sheetView></sheetViews>"u8);
    }

    private static bool HasAttribute(List<KeyValuePair<string, string>>? attributes, string name)
    {
        if (attributes is null)
            return false;

        foreach ((string attribute, _) in attributes)
        {
            if (attribute == name)
                return true;
        }

        return false;
    }

    private void WriteFrozenPane(int rows, int columns)
    {
        _xml.WriteRaw("<pane"u8);
        if (columns > 0)
        {
            _xml.WriteRaw(" xSplit=\""u8);
            _xml.WriteInt32(columns);
            _xml.WriteRaw("\""u8);
        }

        if (rows > 0)
        {
            _xml.WriteRaw(" ySplit=\""u8);
            _xml.WriteInt32(rows);
            _xml.WriteRaw("\""u8);
        }

        _xml.WriteRaw(" topLeftCell=\""u8);
        _xml.WriteFormattable(new CellRef(rows + 1, columns + 1));
        _xml.WriteRaw("\" activePane=\""u8);
        _xml.WriteRaw((rows, columns) switch
        {
            ( > 0, > 0) => "bottomRight"u8,
            ( > 0, _) => "bottomLeft"u8,
            _ => "topRight"u8,
        });
        _xml.WriteRaw("\" state=\"frozen\"/>"u8);
    }

    /// <summary>Writes the column definitions, merging runs of identical neighbours into spans.</summary>
    private void WriteColumns(SheetOptions options)
    {
        if (options.Columns is not { Count: > 0 } columns)
            return;

        List<KeyValuePair<int, ColumnProperties>> ordered = [.. columns.OrderBy(static pair => pair.Key)];
        _xml.WriteRaw("<cols>"u8);
        for (int i = 0; i < ordered.Count;)
        {
            (int first, ColumnProperties properties) = ordered[i];
            int last = first;
            while (++i < ordered.Count && ordered[i].Key == last + 1 && ordered[i].Value == properties)
                last = ordered[i].Key;

            WriteColumnSpan(first, last, properties);
        }

        _xml.WriteRaw("</cols>"u8);
    }

    private void WriteColumnSpan(int first, int last, ColumnProperties properties)
    {
        _xml.WriteRaw("<col min=\""u8);
        _xml.WriteInt32(first);
        _xml.WriteRaw("\" max=\""u8);
        _xml.WriteInt32(last);
        if (properties.Width is { } width)
        {
            // Excel refuses to open a file whose column is wider than 255 characters.
            _xml.WriteRaw("\" width=\""u8);
            _xml.WriteDouble(Math.Min(width, SheetLimits.MaxColumnWidth));
            _xml.WriteRaw("\" customWidth=\"1"u8);
        }

        if (properties.Hidden)
            _xml.WriteRaw("\" hidden=\"1"u8);
        if (properties.OutlineLevel > 0)
        {
            _xml.WriteRaw("\" outlineLevel=\""u8);
            _xml.WriteInt32(properties.OutlineLevel);
        }

        if (properties.Collapsed)
            _xml.WriteRaw("\" collapsed=\"1"u8);
        if (properties.BestFit)
            _xml.WriteRaw("\" bestFit=\"1"u8);
        if (properties.Style != StyleId.Default)
        {
            _xml.WriteRaw("\" style=\""u8);
            _xml.WriteInt32(properties.Style.Index);
        }

        _xml.WriteRaw("\"/>"u8);
    }

    /// <summary>
    /// Writes the auto filter. The criteria of a loaded file come back verbatim while the
    /// range is untouched; once it moves, only the dropdowns are placed, because a filter
    /// condition on a column that is no longer there would hide the wrong rows.
    /// </summary>
    private void WriteAutoFilter(WorksheetMarkup? markup)
    {
        if (_options.AutoFilterRange is not { } range)
            return;

        if (markup?.AutoFilter is { } preserved && markup.AutoFilterRange == range)
        {
            RawXml.Write(_xml, preserved);
            return;
        }

        _xml.WriteRaw("<autoFilter ref=\""u8);
        _xml.WriteFormattable(range);
        _xml.WriteRaw("\"/>"u8);
    }

    /// <summary>
    /// Points the sheet at the drawing holding its pictures. A sheet whose pictures were edited
    /// gets a freshly generated part; an untouched one keeps the element it arrived with, and
    /// with it the charts and shapes that shared the original drawing.
    /// </summary>
    private void WriteDrawing(WorksheetMarkup? markup)
    {
        if (!_options.PicturesChanged)
        {
            RawXml.Write(_xml, markup?.Drawing);
            return;
        }

        if (_options.DrawingRelationshipId is not { } id)
            return;

        _xml.WriteRaw("<drawing r:id=\""u8);
        _xml.WriteAttributeText(id);
        _xml.WriteRaw("\"/>"u8);
    }

    /// <summary>
    /// Points the sheet at the VML holding its note boxes. A sheet whose notes were edited
    /// gets a freshly generated drawing; an untouched one keeps the element it arrived with.
    /// </summary>
    private void WriteLegacyDrawing(WorksheetMarkup? markup)
    {
        if (!_options.CommentsChanged)
        {
            RawXml.Write(_xml, markup?.LegacyDrawing);
            return;
        }

        if (_options.LegacyDrawingRelationshipId is not { } id)
            return;

        _xml.WriteRaw("<legacyDrawing r:id=\""u8);
        _xml.WriteAttributeText(id);
        _xml.WriteRaw("\"/>"u8);
    }

    private void WriteTableParts(WorksheetMarkup? markup)
    {
        if (!_options.TablesChanged)
        {
            RawXml.Write(_xml, markup?.TableParts);
            return;
        }

        if (_options.TableRelationshipIds is not { Count: > 0 } ids)
            return;

        _xml.WriteRaw("<tableParts count=\""u8);
        _xml.WriteInt32(ids.Count);
        _xml.WriteRaw("\">"u8);
        foreach (string id in ids)
        {
            _xml.WriteRaw("<tablePart r:id=\""u8);
            _xml.WriteAttributeText(id);
            _xml.WriteRaw("\"/>"u8);
        }

        _xml.WriteRaw("</tableParts>"u8);
    }

    /// <summary>
    /// Writes the print settings. An untouched sheet gets its original elements back, keeping
    /// the printer settings link and anything else the model does not describe.
    /// </summary>
    private void WritePageSetup(WorksheetMarkup? markup)
    {
        if (_options.PageSetup is not { IsModified: true } setup)
        {
            RawXml.Write(_xml, markup?.PageSetup);
            return;
        }

        PageSetupWriter.Write(_xml, setup, markup);
    }

    /// <summary>
    /// Writes the conditional formatting blocks. An untouched sheet gets its original blocks
    /// back, which keeps the extension-only rule types Excel writes twice.
    /// </summary>
    private void WriteConditionalFormats(WorksheetMarkup? markup)
    {
        if (!_options.ConditionalFormatsChanged)
        {
            RawXml.Write(_xml, markup?.ConditionalFormats);
            return;
        }

        if (_options.ConditionalFormats is not { Count: > 0 } rules || _options.ConditionalStyleIds is not { } styleIds)
            return;

        ConditionalFormatWriter.Write(
            _xml, rules, styleIds, markup?.ExtensionConditionalFormatPriorities);
    }

    /// <summary>
    /// Writes the validation rules. An untouched sheet gets its original element back, which
    /// keeps the x14 extensions Excel attaches for list sources on another sheet.
    /// </summary>
    private void WriteValidations(WorksheetMarkup? markup)
    {
        if (!_options.ValidationsChanged)
        {
            RawXml.Write(_xml, markup?.DataValidations);
            return;
        }

        if (_options.Validations is not { Count: > 0 } validations)
            return;

        _xml.WriteRaw("<dataValidations count=\""u8);
        _xml.WriteInt32(validations.Count);
        _xml.WriteRaw("\">"u8);
        foreach (DataValidation validation in validations)
            WriteValidation(validation);
        _xml.WriteRaw("</dataValidations>"u8);
    }

    private void WriteValidation(DataValidation validation)
    {
        _xml.WriteRaw("<dataValidation"u8);
        if (validation.Type != DataValidationType.None)
        {
            _xml.WriteRaw(" type=\""u8);
            _xml.WriteRaw(validation.Type switch
            {
                DataValidationType.Whole => "whole"u8,
                DataValidationType.Decimal => "decimal"u8,
                DataValidationType.List => "list"u8,
                DataValidationType.Date => "date"u8,
                DataValidationType.Time => "time"u8,
                DataValidationType.TextLength => "textLength"u8,
                _ => "custom"u8,
            });
            _xml.WriteRaw("\""u8);
        }

        if (validation.Operator != DataValidationOperator.Between)
        {
            _xml.WriteRaw(" operator=\""u8);
            _xml.WriteRaw(validation.Operator switch
            {
                DataValidationOperator.NotBetween => "notBetween"u8,
                DataValidationOperator.Equal => "equal"u8,
                DataValidationOperator.NotEqual => "notEqual"u8,
                DataValidationOperator.GreaterThan => "greaterThan"u8,
                DataValidationOperator.LessThan => "lessThan"u8,
                DataValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual"u8,
                _ => "lessThanOrEqual"u8,
            });
            _xml.WriteRaw("\""u8);
        }

        if (validation.ErrorStyle != DataValidationErrorStyle.Stop)
        {
            _xml.WriteRaw(" errorStyle=\""u8);
            _xml.WriteRaw(validation.ErrorStyle == DataValidationErrorStyle.Warning ? "warning"u8 : "information"u8);
            _xml.WriteRaw("\""u8);
        }

        if (validation.AllowBlank)
            _xml.WriteRaw(" allowBlank=\"1\""u8);
        // The file attribute suppresses the arrow, so it is the negation of the model's flag.
        if (!validation.ShowDropDown)
            _xml.WriteRaw(" showDropDown=\"1\""u8);
        if (validation.ShowInputMessage)
            _xml.WriteRaw(" showInputMessage=\"1\""u8);
        if (validation.ShowErrorMessage)
            _xml.WriteRaw(" showErrorMessage=\"1\""u8);

        WriteOptionalAttribute(" promptTitle=\""u8, validation.PromptTitle);
        WriteOptionalAttribute(" prompt=\""u8, validation.Prompt);
        WriteOptionalAttribute(" errorTitle=\""u8, validation.ErrorTitle);
        WriteOptionalAttribute(" error=\""u8, validation.Error);

        _xml.WriteRaw(" sqref=\""u8);
        for (int i = 0; i < validation.Ranges.Count; i++)
        {
            if (i > 0)
                _xml.WriteRaw(" "u8);
            _xml.WriteFormattable(validation.Ranges[i]);
        }

        if (validation.Formula1 is null && validation.Formula2 is null)
        {
            _xml.WriteRaw("\"/>"u8);
            return;
        }

        _xml.WriteRaw("\">"u8);
        WriteFormula("<formula1>"u8, "</formula1>"u8, validation.Formula1);
        WriteFormula("<formula2>"u8, "</formula2>"u8, validation.Formula2);
        _xml.WriteRaw("</dataValidation>"u8);
    }

    private void WriteOptionalAttribute(scoped ReadOnlySpan<byte> prefix, string? value)
    {
        if (value is null)
            return;
        _xml.WriteRaw(prefix);
        _xml.WriteAttributeText(value);
        _xml.WriteRaw("\""u8);
    }

    private void WriteFormula(scoped ReadOnlySpan<byte> open, scoped ReadOnlySpan<byte> close, string? formula)
    {
        if (formula is null)
            return;
        _xml.WriteRaw(open);
        _xml.WriteText(FutureFunctions.QualifyAll(formula)!);
        _xml.WriteRaw(close);
    }

    private void WriteHyperlinks()
    {
        if (_options.Hyperlinks is not { Count: > 0 } links)
            return;

        _xml.WriteRaw("<hyperlinks>"u8);
        foreach (WrittenHyperlink link in links)
        {
            _xml.WriteRaw("<hyperlink ref=\""u8);
            _xml.WriteFormattable(link.Range);
            if (link.RelationshipId is { } id)
            {
                _xml.WriteRaw("\" r:id=\""u8);
                _xml.WriteAttributeText(id);
            }

            if (link.Location is { } location)
            {
                _xml.WriteRaw("\" location=\""u8);
                _xml.WriteAttributeText(location);
            }

            if (link.Tooltip is { } tooltip)
            {
                _xml.WriteRaw("\" tooltip=\""u8);
                _xml.WriteAttributeText(tooltip);
            }

            if (link.Display is { } display)
            {
                _xml.WriteRaw("\" display=\""u8);
                _xml.WriteAttributeText(display);
            }

            _xml.WriteRaw("\"/>"u8);
        }

        _xml.WriteRaw("</hyperlinks>"u8);
    }

    private void WriteCell(int row, int column, in CellValue value, int styleIndex, string? formula)
    {
        if (formula is null && value.IsBlank && styleIndex == 0)
            return;

        // Unstyled temporal values get the fixed date/duration formats so Excel renders them.
        int effectiveStyle = styleIndex != 0 ? styleIndex : value.Kind switch
        {
            CellValueKind.DateTime => StylePool.DateTimeStyleIndex,
            CellValueKind.TimeSpan => StylePool.TimeSpanStyleIndex,
            _ => 0,
        };

        var reference = new CellRef(row, column);
        _xml.WriteRaw("<c r=\""u8);
        _xml.WriteFormattable(reference);
        _xml.WriteRaw("\""u8);
        if (effectiveStyle != 0)
        {
            _xml.WriteRaw(" s=\""u8);
            _xml.WriteInt32(effectiveStyle);
            _xml.WriteRaw("\""u8);
        }

        // The cm and vm indexes tying a cell to the metadata part behind rich values.
        if (_cellMarkup is not null && _cellMarkup.TryGetValue(reference, out List<KeyValuePair<string, string>>? extra))
            RawXml.WriteAttributes(_xml, extra);

        FormulaMarkup? markup = _formulaMarkup?.TryGetValue(reference, out FormulaMarkup found) == true ? found : null;

        // A data table has attributes but no text, so the cell still needs its formula element
        // even though the model holds no formula for it.
        if (formula is not null || markup is { WritesOwnElement: true })
        {
            WriteFormulaCell(in value, formula, reference, markup);
            return;
        }

        switch (value.Kind)
        {
            case CellValueKind.Number:
                _xml.WriteRaw("><v>"u8);
                _xml.WriteDouble(value.GetNumber());
                _xml.WriteRaw("</v></c>"u8);
                break;

            case CellValueKind.Text when _owner.Options.UseInlineStrings:
                _xml.WriteRaw(" t=\"inlineStr\"><is>"u8);
                WriteTextRun(_xml, value.GetText());
                _xml.WriteRaw("</is></c>"u8);
                break;

            case CellValueKind.Text:
                _xml.WriteRaw(" t=\"s\"><v>"u8);
                string text = value.GetText();
                string? richMarkup = _richStrings?.TryGetValue(reference, out RichStringMarkup rich) == true &&
                    rich.Text == text
                    ? rich.Markup
                    : null;
                _xml.WriteInt32(_owner.SharedStrings.GetOrAdd(text, richMarkup));
                _xml.WriteRaw("</v></c>"u8);
                break;

            case CellValueKind.Boolean:
                _xml.WriteRaw(" t=\"b\"><v>"u8);
                _xml.WriteRaw(value.GetBoolean() ? "1"u8 : "0"u8);
                _xml.WriteRaw("</v></c>"u8);
                break;

            case CellValueKind.DateTime or CellValueKind.TimeSpan:
                _xml.WriteRaw("><v>"u8);
                _xml.WriteDouble(value.ToSerialNumber(_owner.Options.DateSystem));
                _xml.WriteRaw("</v></c>"u8);
                break;

            case CellValueKind.Error:
                _xml.WriteRaw(" t=\"e\"><v>"u8);
                _xml.WriteText(value.GetError().ToText());
                _xml.WriteRaw("</v></c>"u8);
                break;

            default:
                _xml.WriteRaw("/>"u8);
                break;
        }
    }

    /// <summary>Writes <c>&lt;f&gt;</c> plus the cached value with the type attribute the value kind demands.</summary>
    private void WriteFormulaCell(in CellValue value, string? formula, CellRef cell, FormulaMarkup? markup)
    {
        _owner.OnFormulaWritten();
        switch (value.Kind)
        {
            case CellValueKind.Text:
                _xml.WriteRaw(" t=\"str\""u8);
                break;
            case CellValueKind.Boolean:
                _xml.WriteRaw(" t=\"b\""u8);
                break;
            case CellValueKind.Error:
                _xml.WriteRaw(" t=\"e\""u8);
                break;
            default:
                break;
        }

        _xml.WriteRaw(">"u8);
        WriteFormulaElement(formula, cell, markup);

        switch (value.Kind)
        {
            case CellValueKind.Number:
                _xml.WriteRaw("<v>"u8);
                _xml.WriteDouble(value.GetNumber());
                _xml.WriteRaw("</v>"u8);
                break;
            case CellValueKind.Text:
                _xml.WriteRaw("<v>"u8);
                _xml.WriteText(value.GetText());
                _xml.WriteRaw("</v>"u8);
                break;
            case CellValueKind.Boolean:
                _xml.WriteRaw(value.GetBoolean() ? "<v>1</v>"u8 : "<v>0</v>"u8);
                break;
            case CellValueKind.DateTime or CellValueKind.TimeSpan:
                _xml.WriteRaw("<v>"u8);
                _xml.WriteDouble(value.ToSerialNumber(_owner.Options.DateSystem));
                _xml.WriteRaw("</v>"u8);
                break;
            case CellValueKind.Error:
                _xml.WriteRaw("<v>"u8);
                _xml.WriteText(value.GetError().ToText());
                _xml.WriteRaw("</v>"u8);
                break;
            default:
                break;
        }

        _xml.WriteRaw("</c>"u8);
    }

    /// <summary>Writes the formula element, restoring the array marker for registered masters.</summary>
    private void WriteFormulaElement(string? formula, CellRef cell, FormulaMarkup? markup)
    {
        if (markup is { WritesOwnElement: true } own)
        {
            _xml.WriteRaw("<f"u8);
            RawXml.WriteAttributes(_xml, own.Attributes);
            _xml.WriteRaw("/>"u8);
            return;
        }

        _xml.WriteRaw("<f"u8);
        if (_arrayFormulaRanges is not null && _arrayFormulaRanges.TryGetValue(cell, out RangeRef range))
        {
            _xml.WriteRaw(" t=\"array\" ref=\""u8);
            _xml.WriteFormattable(range);
            _xml.WriteRaw("\""u8);
        }

        RawXml.WriteAttributes(_xml, markup?.Attributes);
        _xml.WriteRaw(">"u8);

        // A function the standard never listed has to be spelled _xlfn.NAME here, or Excel
        // reads the file back as #NAME? even though it implements the function.
        _xml.WriteText(FutureFunctions.QualifyAll(formula) ?? "");
        _xml.WriteRaw("</f>"u8);
    }

    private void WriteMergeCells()
    {
        if (_merges.Count == 0)
            return;

        _xml.WriteRaw("<mergeCells count=\""u8);
        _xml.WriteInt32(_merges.Count);
        _xml.WriteRaw("\">"u8);
        foreach (RangeRef merge in _merges)
        {
            _xml.WriteRaw("<mergeCell ref=\""u8);
            _xml.WriteFormattable(merge);
            _xml.WriteRaw("\"/>"u8);
        }

        _xml.WriteRaw("</mergeCells>"u8);
    }

    /// <summary>
    /// Writes one shared-string item: <c>&lt;si&gt;&lt;t&gt;…&lt;/t&gt;&lt;/si&gt;</c>, or the
    /// original runs when the string arrived with formatting and nobody changed its text.
    /// </summary>
    internal static void WriteSharedStringItem(Utf8XmlWriter xml, string text, string? richMarkup)
    {
        xml.WriteRaw("<si>"u8);
        if (richMarkup is null)
            WriteTextRun(xml, text);
        else
            RawXml.Write(xml, richMarkup);
        xml.WriteRaw("</si>"u8);
    }

    private static void WriteTextRun(Utf8XmlWriter xml, string text)
    {
        if (NeedsSpacePreserve(text))
            xml.WriteRaw("<t xml:space=\"preserve\">"u8);
        else
            xml.WriteRaw("<t>"u8);
        xml.WriteText(text);
        xml.WriteRaw("</t>"u8);
    }

    internal static bool NeedsSpacePreserve(string text) =>
        text.Length > 0 &&
        (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]) || text.Contains('\n'));

    private static bool AllBlank(scoped ReadOnlySpan<CellValue> cells)
    {
        foreach (ref readonly CellValue cell in cells)
        {
            if (!cell.IsBlank)
                return false;
        }

        return true;
    }
}
