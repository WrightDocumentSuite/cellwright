using System.Globalization;
using System.Text;
using System.Xml;
using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Xml;

namespace Cellwright;

public sealed partial class XlsxReader
{
    /// <summary>
    /// Reads everything the DOM cares about from one worksheet part — sheet properties, views,
    /// column and row definitions, cells (with styles and formulas), merges and hyperlinks —
    /// into a sink, capturing the rest of the part as markup to replay on save.
    /// </summary>
    internal async ValueTask ReadSheetContentAsync(int sheetIndex, ISheetContentSink sink, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await EnsureSharedStringsAsync(cancellationToken).ConfigureAwait(false);
        await EnsureStylesAsync(cancellationToken).ConfigureAwait(false);

        var capture = new SheetCapture(sink);
        int lastRowNumber = 0;
        Stream stream = await _package.OpenPartAsync(_sheetParts[sheetIndex], cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            using var xml = XmlReader.Create(stream, XmlDefaults.AsyncReaderSettings);
            bool advance = true;
            while (true)
            {
                if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                    break;
                advance = true;

                if (xml.NodeType != XmlNodeType.Element)
                    continue;

                switch (xml.LocalName)
                {
                    case "worksheet":
                        capture.Markup.RootAttributes = RawXml.ReadAttributes(xml, "xmlns", "xmlns:r");
                        break;

                    case "sheetPr":
                        await ParseSheetPropertiesAsync(xml, capture).ConfigureAwait(false);
                        break;

                    case "sheetView":
                        await ParseSheetViewAsync(xml, capture).ConfigureAwait(false);
                        break;

                    case "pane":
                        ParsePane(xml, sink);
                        break;

                    case "col":
                        ParseColumnDefinition(xml, sink);
                        break;

                    case "row":
                        cancellationToken.ThrowIfCancellationRequested();
                        lastRowNumber = await ParseDomRowAsync(xml, sink, lastRowNumber).ConfigureAwait(false);
                        break;

                    case "mergeCell":
                        if (RangeRef.TryParse(xml.GetAttribute("ref"), null, out RangeRef merge))
                            sink.OnMerge(merge);
                        break;

                    case "autoFilter":
                        await capture.CaptureAutoFilterAsync(xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    case "hyperlink":
                        ParseHyperlink(xml, sink);
                        break;

                    case "conditionalFormatting":
                        await capture.CaptureConditionalFormatsAsync(xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    case "dataValidations":
                        await capture.CaptureValidationsAsync(xml).ConfigureAwait(false);
                        advance = false;
                        break;

                    // Containers whose children are handled above, plus the stale dimension hint.
                    case "sheetViews" or "cols" or "sheetData" or "mergeCells" or "hyperlinks" or "dimension":
                        break;

                    default:
                        await capture.AppendAsync(xml).ConfigureAwait(false);
                        advance = false;
                        break;
                }
            }
        }

        capture.Finish();
        sink.OnMarkup(capture.Markup);
    }

    /// <summary>
    /// Reads <c>sheetPr</c>: the tab color becomes a model property, everything else is kept
    /// verbatim so that outline direction and page-setup flags survive.
    /// </summary>
    private static async ValueTask ParseSheetPropertiesAsync(XmlReader xml, SheetCapture capture)
    {
        capture.Markup.SheetProperties = RawXml.ReadAttributes(xml);
        if (xml.IsEmptyElement)
            return;

        StringBuilder? children = null;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "sheetPr")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "tabColor")
            {
                capture.Sink.OnTabColor(ParseRgbColor(xml));
            }
            else if (xml.LocalName == "pageSetUpPr")
            {
                PageSetupPart.ReadPageSetupProperties(xml, capture.Sink.PageSetup);
                capture.Markup.PageSetupProperties = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                advance = false;
            }
            else
            {
                children = await RawXml.AppendOuterXmlAsync(children, xml).ConfigureAwait(false);
                advance = false;
            }
        }

        capture.Markup.SheetPropertyChildren = children?.ToString();
    }

    /// <summary>
    /// Reads the first <c>sheetView</c>: grid lines and zoom become model properties, the pane
    /// is reported to the sink, and the rest (selections, tab state) is kept verbatim.
    /// Additional views belong to other workbook views and are dropped.
    /// </summary>
    private static async ValueTask ParseSheetViewAsync(XmlReader xml, SheetCapture capture)
    {
        if (capture.SawSheetView)
        {
            if (!xml.IsEmptyElement)
                await xml.SkipAsync().ConfigureAwait(false);
            return;
        }

        capture.SawSheetView = true;
        if (xml.GetAttribute("showGridLines") is "0" or "false")
            capture.Sink.OnGridLinesHidden();
        if (int.TryParse(xml.GetAttribute("zoomScale"), NumberStyles.None, CultureInfo.InvariantCulture, out int zoom))
            capture.Sink.OnZoomScale(zoom);

        capture.Markup.SheetViewAttributes = RawXml.ReadAttributes(xml, "showGridLines", "zoomScale");
        if (xml.IsEmptyElement)
            return;

        StringBuilder? children = null;
        bool advance = true;
        while (true)
        {
            if (advance && !await xml.ReadAsync().ConfigureAwait(false))
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "sheetView")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "pane")
            {
                ParsePane(xml, capture.Sink);
            }
            else
            {
                children = await RawXml.AppendOuterXmlAsync(children, xml).ConfigureAwait(false);
                advance = false;
            }
        }

        capture.Markup.SheetViewChildren = children?.ToString();
    }

    private static ExcelColor ParseRgbColor(XmlReader xml) =>
        xml.GetAttribute("rgb") is { } rgb && ExcelColor.TryParseHex(rgb, out ExcelColor color)
            ? color
            : ExcelColor.Auto;

    private static void ParsePane(XmlReader xml, ISheetContentSink sink)
    {
        if (xml.GetAttribute("state") is not ("frozen" or "frozenSplit"))
            return;

        int rows = ParseSplit(xml.GetAttribute("ySplit"));
        int columns = ParseSplit(xml.GetAttribute("xSplit"));
        if (rows > 0 || columns > 0)
            sink.OnFrozenPane(rows, columns);

        static int ParseSplit(string? value) =>
            double.TryParse(value, CultureInfo.InvariantCulture, out double split) ? (int)split : 0;
    }

    /// <summary>
    /// An unsigned attribute of the file. Always invariant: what a package holds is the same
    /// number wherever it is opened, and the culture of the machine reading it has no say.
    /// </summary>
    private static int Integer(XmlReader xml, string attribute) =>
        int.TryParse(xml.GetAttribute(attribute), NumberStyles.None, CultureInfo.InvariantCulture, out int value)
            ? value
            : 0;

    private static void ParseColumnDefinition(XmlReader xml, ISheetContentSink sink)
    {
        if (!int.TryParse(xml.GetAttribute("min"), NumberStyles.None, CultureInfo.InvariantCulture, out int min) ||
            !int.TryParse(xml.GetAttribute("max"), NumberStyles.None, CultureInfo.InvariantCulture, out int max))
        {
            return;
        }

        double? width = double.TryParse(xml.GetAttribute("width"), CultureInfo.InvariantCulture, out double parsed)
            && xml.GetAttribute("customWidth") is "1" or "true"
            ? parsed
            : null;

        var definition = new ColumnDefinition(
            Column: 0,
            Width: width,
            Hidden: xml.GetAttribute("hidden") is "1" or "true",
            OutlineLevel: Integer(xml, "outlineLevel"),
            Collapsed: xml.GetAttribute("collapsed") is "1" or "true",
            RawStyleIndex: Integer(xml, "style"),
            BestFit: xml.GetAttribute("bestFit") is "1" or "true");
        if (definition.IsEmpty)
            return;

        min = Math.Max(min, 1);
        max = Math.Min(max, SheetLimits.MaxColumn);
        for (int column = min; column <= max; column++)
            sink.OnColumn(definition with { Column = column });
    }

    /// <summary>Reads one <c>dataValidation</c> element, including its two formula children.</summary>
    private static void ParseDataValidation(XmlReader fragment, ISheetContentSink sink)
    {
        if (fragment.GetAttribute("sqref") is not { Length: > 0 } sqref)
            return;

        IReadOnlyList<RangeRef> ranges;
        try
        {
            ranges = DataValidation.ParseRanges(sqref);
        }
        catch (FormatException)
        {
            return;
        }

        DataValidationType type = fragment.GetAttribute("type") switch
        {
            "whole" => DataValidationType.Whole,
            "decimal" => DataValidationType.Decimal,
            "list" => DataValidationType.List,
            "date" => DataValidationType.Date,
            "time" => DataValidationType.Time,
            "textLength" => DataValidationType.TextLength,
            "custom" => DataValidationType.Custom,
            _ => DataValidationType.None,
        };

        DataValidationOperator comparison = fragment.GetAttribute("operator") switch
        {
            "notBetween" => DataValidationOperator.NotBetween,
            "equal" => DataValidationOperator.Equal,
            "notEqual" => DataValidationOperator.NotEqual,
            "greaterThan" => DataValidationOperator.GreaterThan,
            "lessThan" => DataValidationOperator.LessThan,
            "greaterThanOrEqual" => DataValidationOperator.GreaterThanOrEqual,
            "lessThanOrEqual" => DataValidationOperator.LessThanOrEqual,
            _ => DataValidationOperator.Between,
        };

        DataValidationErrorStyle errorStyle = fragment.GetAttribute("errorStyle") switch
        {
            "warning" => DataValidationErrorStyle.Warning,
            "information" => DataValidationErrorStyle.Information,
            _ => DataValidationErrorStyle.Stop,
        };

        bool allowBlank = IsTrue(fragment, "allowBlank");
        bool showDropDown = !IsTrue(fragment, "showDropDown"); // The attribute means "suppress".
        bool showInput = IsTrue(fragment, "showInputMessage");
        bool showError = IsTrue(fragment, "showErrorMessage");
        string? promptTitle = fragment.GetAttribute("promptTitle");
        string? prompt = fragment.GetAttribute("prompt");
        string? errorTitle = fragment.GetAttribute("errorTitle");
        string? error = fragment.GetAttribute("error");
        string? formula1 = ReadFormulas(fragment, out string? formula2);

        sink.OnDataValidation(new DataValidation
        {
            Ranges = ranges,
            Type = type,
            Operator = comparison,
            AllowBlank = allowBlank,
            ShowDropDown = showDropDown,
            ShowInputMessage = showInput,
            ShowErrorMessage = showError,
            PromptTitle = promptTitle,
            Prompt = prompt,
            ErrorTitle = errorTitle,
            Error = error,
            ErrorStyle = errorStyle,
            Formula1 = formula1,
            Formula2 = formula2,
        });

        static bool IsTrue(XmlReader xml, string name) => xml.GetAttribute(name) is "1" or "true";
    }

    /// <summary>Reads both formula children of the current element in one pass over its subtree.</summary>
    private static string? ReadFormulas(XmlReader xml, out string? formula2)
    {
        string? formula1 = null;
        formula2 = null;
        if (xml.IsEmptyElement)
            return null;

        using XmlReader subtree = xml.ReadSubtree();
        subtree.Read();
        bool advance = true;
        while (true)
        {
            // Reading an element's text steps past its end tag, so advancing again here would
            // walk straight over the sibling formula.
            if (advance && !subtree.Read())
                break;
            advance = true;

            if (subtree.NodeType != XmlNodeType.Element)
                continue;
            if (subtree.LocalName == "formula1")
            {
                formula1 = FutureFunctions.UnqualifyAll(subtree.ReadElementContentAsString());
                advance = false;
            }
            else if (subtree.LocalName == "formula2")
            {
                formula2 = FutureFunctions.UnqualifyAll(subtree.ReadElementContentAsString());
                advance = false;
            }
        }

        return formula1;
    }

    private static void ParseHyperlink(XmlReader xml, ISheetContentSink sink)
    {
        if (!RangeRef.TryParse(xml.GetAttribute("ref"), null, out RangeRef range))
            return;

        sink.OnHyperlink(
            range,
            RelationshipId(xml),
            xml.GetAttribute("location"),
            xml.GetAttribute("tooltip"),
            xml.GetAttribute("display"));
    }

    /// <summary>Parses one row element for the DOM, reporting cells through the sink.</summary>
    private async ValueTask<int> ParseDomRowAsync(XmlReader xml, ISheetContentSink sink, int lastRowNumber)
    {
        int rowNumber = int.TryParse(xml.GetAttribute("r"), NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : lastRowNumber + 1;

        double? height = xml.GetAttribute("customHeight") is "1" or "true" &&
            double.TryParse(xml.GetAttribute("ht"), CultureInfo.InvariantCulture, out double parsedHeight)
            ? parsedHeight
            : null;

        // The style only counts when customFormat says the row was formatted on purpose;
        // Excel leaves a stale index behind otherwise.
        var definition = new RowDefinition(
            Row: rowNumber,
            Height: height,
            Hidden: xml.GetAttribute("hidden") is "1" or "true",
            OutlineLevel: Integer(xml, "outlineLevel"),
            Collapsed: xml.GetAttribute("collapsed") is "1" or "true",
            RawStyleIndex: xml.GetAttribute("customFormat") is "1" or "true" ? Integer(xml, "s") : 0,
            ThickTop: xml.GetAttribute("thickTop") is "1" or "true",
            ThickBottom: xml.GetAttribute("thickBot") is "1" or "true");
        if (!definition.IsEmpty)
            sink.OnRow(definition);

        if (xml.IsEmptyElement)
            return rowNumber;

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

            string? reference = xml.GetAttribute("r");
            int column = ResolveColumn(reference, lastColumn, rowNumber);
            lastColumn = column;
            if (column > SheetLimits.MaxColumn)
                continue;

            string? styleAttribute = xml.GetAttribute("s");
            string? type = xml.GetAttribute("t");
            int styleIndex = int.TryParse(styleAttribute, NumberStyles.None, CultureInfo.InvariantCulture, out int style) ? style : 0;
            var cell = new CellRef(rowNumber, column);

            // Counting first keeps the common cell — which carries nothing but r, s and t —
            // from walking its own attribute list a second time.
            int modelled = (reference is null ? 0 : 1) + (styleAttribute is null ? 0 : 1) + (type is null ? 0 : 1);
            if (xml.AttributeCount > modelled && RawXml.ReadAttributes(xml, "r", "s", "t") is { } extra)
                sink.OnCellMarkup(cell, extra);

            CellValue value = CellValue.Blank;
            ParsedFormula formula = default;
            int? sharedStringIndex = null;
            if (!xml.IsEmptyElement)
                (value, formula, sharedStringIndex) = await ParseCellContentAsync(xml, type, styleIndex, cell, captureFormula: true).ConfigureAwait(false);

            if (sharedStringIndex is int index && _richStrings![index] is string rich)
                sink.OnRichString(cell, value.GetText(), rich);

            string? formulaText = ResolveFormulaText(in formula, cell, sink);
            if (!value.IsBlank || styleIndex != 0 || formulaText is not null)
                sink.OnCell(cell, in value, styleIndex, formulaText);
        }

        return rowNumber;
    }

    /// <summary>
    /// Turns the raw <c>f</c> element into the text stored on the cell and reports shared and
    /// array groups to the sink. Shared members carry no text of their own: the group is what
    /// lets the calculation engine reconstruct them later.
    /// </summary>
    private string? ResolveFormulaText(in ParsedFormula formula, CellRef cell, ISheetContentSink sink)
    {
        if (formula.Markup is { } markup)
            sink.OnFormulaMarkup(cell, markup);

        if (formula.Text is null)
        {
            if (formula is { IsShared: true, SharedIndex: { } memberIndex })
                sink.OnSharedFormulaMember(cell, memberIndex);
            return null;
        }

        // The file spells a post-2007 function as _xlfn.NAME; that prefix belongs to the format
        // and not to the formula, so the model keeps the name a person would type. Saving puts
        // it back.
        string text = FutureFunctions.UnqualifyAll(formula.Text)!;

        if (formula.IsShared && formula.SharedIndex is { } masterIndex)
        {
            sink.OnFormulaGroup(
                new FormulaGroup(FormulaGroupKind.Shared, cell, formula.Range ?? new RangeRef(cell), text),
                masterIndex);
        }
        else if (formula.IsArray)
        {
            sink.OnFormulaGroup(
                new FormulaGroup(FormulaGroupKind.Array, cell, formula.Range ?? new RangeRef(cell), text),
                sharedIndex: null);
        }

        return text;
    }

    /// <summary>
    /// Routes each unmodelled worksheet element into the schema slot it has to be replayed
    /// from. Elements arrive in document order, which is schema order in any readable file.
    /// </summary>
    private sealed class SheetCapture(ISheetContentSink sink)
    {
        private StringBuilder? _buffer;
        private Slot _slot = Slot.BeforeAutoFilter;

        private enum Slot
        {
            BeforeAutoFilter,
            AfterProtection,
            BeforeMerges,
            AfterMerges,
            PrintSetup,
            Drawing,
            AfterLegacyDrawing,
        }

        public ISheetContentSink Sink => sink;

        public WorksheetMarkup Markup { get; } = new();

        public bool SawSheetView { get; set; }

        /// <summary>
        /// Keeps the filter's range in the model and its criteria as markup. The criteria are
        /// replayed only while the range still matches, since a moved filter cannot keep them.
        /// </summary>
        public async ValueTask CaptureAutoFilterAsync(XmlReader xml)
        {
            Flush();
            bool hasRange = RangeRef.TryParse(xml.GetAttribute("ref"), null, out RangeRef range);
            Markup.AutoFilter = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
            _slot = Slot.BeforeMerges;
            if (!hasRange)
                return;

            Markup.AutoFilterRange = range;

            // A filterColumn child is the criteria; without one the dropdowns show everything.
            sink.OnAutoFilter(range, Markup.AutoFilter.Contains("filterColumn", StringComparison.Ordinal));
        }

        /// <summary>
        /// Reads one conditional formatting block into the model and keeps it as markup. A
        /// sheet can carry several blocks, so they accumulate into one fragment replayed in
        /// document order.
        /// </summary>
        public async ValueTask CaptureConditionalFormatsAsync(XmlReader xml)
        {
            Flush();
            _slot = Slot.PrintSetup;
            string block = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
            Markup.ConditionalFormats += block;
            ConditionalFormatPart.Read(block, sink);
        }

        /// <summary>
        /// Reads the validation rules into the model and keeps the element as markup: the
        /// x14 extensions Excel adds for cross-sheet list sources hang off it, and they only
        /// stay meaningful while the rules themselves are untouched.
        /// </summary>
        public async ValueTask CaptureValidationsAsync(XmlReader xml)
        {
            Flush();
            _slot = Slot.PrintSetup;
            if (xml.IsEmptyElement)
            {
                await xml.ReadAsync().ConfigureAwait(false);
                return;
            }

            string content = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
            Markup.DataValidations = content;
            using var fragment = XmlReader.Create(new StringReader(content), XmlDefaults.ReaderSettings);
            while (fragment.Read())
            {
                if (fragment.NodeType == XmlNodeType.Element && fragment.LocalName == "dataValidation")
                    ParseDataValidation(fragment, sink);
            }
        }

        public async ValueTask AppendAsync(XmlReader xml)
        {
            switch (xml.LocalName)
            {
                case "sheetFormatPr":
                    Markup.SheetFormatProperties = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    return;

                case "sheetProtection":
                    Flush();
                    _slot = Slot.AfterProtection;
                    ProtectionPart.ReadSheet(xml, sink.Protection);
                    await xml.SkipAsync().ConfigureAwait(false);
                    return;

                case "protectedRanges":
                {
                    Flush();
                    _slot = Slot.AfterProtection;
                    string ranges = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    ProtectionPart.ReadProtectedRanges(ranges, sink.Protection);
                    return;
                }

                // The print settings are modelled, but they still travel as markup so that an
                // untouched sheet keeps whatever the producer wrote down to the attribute.
                case "printOptions" or "pageMargins" or "pageSetup" or "headerFooter" or "rowBreaks" or "colBreaks":
                {
                    Flush();
                    _slot = Slot.PrintSetup;
                    string element = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    Markup.PageSetup += element;
                    PageSetupPart.Read(element, sink.PageSetup, Markup);
                    return;
                }

                case "legacyDrawing":
                    Flush();
                    Markup.LegacyDrawing = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    _slot = Slot.AfterLegacyDrawing;
                    return;
                case "tableParts":
                    Flush();
                    Markup.TableParts = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    return;
                case "extLst":
                    Flush();
                    Markup.Extensions = await xml.ReadOuterXmlAsync().ConfigureAwait(false);
                    return;
                default:
                    break;
            }

            Slot slot = SlotOf(xml.LocalName);
            if (slot != _slot)
            {
                Flush();
                _slot = slot;
            }

            _buffer = await RawXml.AppendOuterXmlAsync(_buffer, xml).ConfigureAwait(false);
        }

        public void Finish() => Flush();

        private static Slot SlotOf(string name) => name switch
        {
            "sheetCalcPr" => Slot.BeforeAutoFilter,
            "scenarios" => Slot.AfterProtection,
            "sortState" or "dataConsolidate" or "customSheetViews" => Slot.BeforeMerges,
            "phoneticPr" => Slot.AfterMerges,
            "drawing" => Slot.Drawing,

            // Everything the schema places after legacyDrawing, including the AlternateContent
            // wrapper Excel puts around form controls and OLE objects.
            "legacyDrawingHF" or "drawingHF" or "picture" or
            "oleObjects" or "controls" or "webPublishItems" or "AlternateContent" => Slot.AfterLegacyDrawing,

            _ => Slot.PrintSetup,
        };

        private void Flush()
        {
            if (_buffer is null)
                return;

            string content = _buffer.ToString();
            _buffer = null;
            switch (_slot)
            {
                case Slot.BeforeAutoFilter:
                    Markup.BeforeAutoFilter += content;
                    break;
                case Slot.AfterProtection:
                    Markup.AfterProtection += content;
                    break;
                case Slot.BeforeMerges:
                    Markup.BeforeMerges += content;
                    break;
                case Slot.AfterMerges:
                    Markup.AfterMerges += content;
                    break;
                case Slot.Drawing:
                    Markup.Drawing += content;
                    break;
                case Slot.AfterLegacyDrawing:
                    Markup.AfterLegacyDrawing += content;
                    break;
                default:
                    Markup.PrintSetup += content;
                    break;
            }
        }
    }
}
