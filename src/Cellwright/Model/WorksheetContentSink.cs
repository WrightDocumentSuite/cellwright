using Cellwright.IO;
using Cellwright.Streaming;
using Cellwright.Styles;

namespace Cellwright.Model;

/// <summary>
/// Routes streamed sheet content into a worksheet's column slices. Shared by every reader that
/// produces a DOM — the xlsx one and the legacy .xls one — so that both end at the same model.
/// </summary>
internal sealed class WorksheetContentSink(
    Worksheet sheet,
    int[] styleMap,
    IReadOnlyList<OpcRelationship> relationships,
    bool preserveMarkup,
    Action<WorkbookWarning> warn,
    DocumentLoadBudgetState? loadBudget = null,
    bool accountText = false) : ISheetContentSink
{
    private readonly Dictionary<int, FormulaGroup> _sharedGroups = [];
    private bool _warnedAboutDataTables;

    /// <summary>The worksheet being filled.</summary>
    public Worksheet Sheet => sheet;

    public void OnFrozenPane(int rows, int columns)
    {
        sheet.FrozenRows = rows;
        sheet.FrozenColumns = columns;
    }

    public void OnTabColor(ExcelColor color) => sheet.TabColor = color;

    public void OnGridLinesHidden() => sheet.ShowGridLines = false;

    public void OnZoomScale(int percent) => sheet.ZoomScale = percent;

    public void OnColumn(in ColumnDefinition definition) =>
        sheet.ColumnTable[definition.Column] = new ColumnProperties
        {
            Width = definition.Width,
            Hidden = definition.Hidden,
            OutlineLevel = definition.OutlineLevel,
            Collapsed = definition.Collapsed,
            Style = MapStyle(definition.RawStyleIndex),
            BestFit = definition.BestFit,
        };

    public void OnRow(in RowDefinition definition) =>
        sheet.RowTable[definition.Row] = new RowProperties
        {
            Height = definition.Height,
            Hidden = definition.Hidden,
            OutlineLevel = definition.OutlineLevel,
            Collapsed = definition.Collapsed,
            Style = MapStyle(definition.RawStyleIndex),
            ThickTop = definition.ThickTop,
            ThickBottom = definition.ThickBottom,
        };

    /// <summary>Turns a styles-part index into a handle in the workbook's pool.</summary>
    private StyleId MapStyle(int rawStyleIndex) =>
        (uint)rawStyleIndex < (uint)styleMap.Length ? new StyleId(styleMap[rawStyleIndex]) : StyleId.Default;

    public void OnHyperlink(RangeRef range, string? relationshipId, string? location, string? tooltip, string? display)
    {
        string? target = relationshipId is null ? null : ResolveTarget(relationshipId);
        if (target is null && location is null)
        {
            warn(new WorkbookWarning("BrokenHyperlink", $"The hyperlink over {range} points nowhere and was dropped."));
            return;
        }

        sheet.Hyperlinks.Add(Hyperlink.Load(range, target, location, tooltip, display));
    }

    /// <summary>
    /// Adds a hyperlink whose target is already resolved, which is how the legacy format states
    /// it: a .xls file has no relationship part to look the target up in.
    /// </summary>
    public void OnResolvedHyperlink(RangeRef range, string? target, string? location, string? tooltip, string? display)
    {
        if (target is null && location is null)
            return;

        sheet.Hyperlinks.Add(Hyperlink.Load(range, target, location, tooltip, display));
    }

    public void OnMarkup(WorksheetMarkup markup)
    {
        if (preserveMarkup)
            sheet.Markup = markup;
    }

    private string? ResolveTarget(string relationshipId)
    {
        foreach (OpcRelationship relationship in relationships)
        {
            if (relationship.Id == relationshipId)
                return relationship.Target;
        }

        return null;
    }

    public void OnCell(CellRef cell, in CellValue value, int rawStyleIndex, string? formula)
    {
        Account(value, formula);
        if (!value.IsBlank)
            sheet.Values.Set(cell.Row, cell.Column, value);
        if (formula is not null)
            sheet.Formulas.Set(cell.Row, cell.Column, formula);

        if (rawStyleIndex == 0)
            return;
        if ((uint)rawStyleIndex < (uint)styleMap.Length)
            sheet.StyleIndexes.Set(cell.Row, cell.Column, styleMap[rawStyleIndex]);
        else
            warn(new WorkbookWarning("BadStyleIndex", $"Style index {rawStyleIndex} is out of range.", cell));
    }

    public void OnRichString(CellRef cell, string text, string markup)
    {
        if (preserveMarkup)
            sheet.RichStringTable[cell] = new RichStringMarkup(text, markup);
    }

    /// <summary>
    /// Sets a cell whose style is already a handle in the workbook's pool. The legacy reader
    /// maps its own style table up front, so there is no second index space to go through.
    /// </summary>
    public void OnPooledCell(CellRef cell, in CellValue value, StyleId style, string? formula)
    {
        Account(value, formula);
        if (!value.IsBlank)
            sheet.Values.Set(cell.Row, cell.Column, value);
        if (formula is not null)
            sheet.Formulas.Set(cell.Row, cell.Column, formula);
        if (style != StyleId.Default)
            sheet.StyleIndexes.Set(cell.Row, cell.Column, style.Index);
    }

    private void Account(in CellValue value, string? formula)
    {
        if (loadBudget is null)
            return;

        loadBudget.AddCell();
        if (!accountText)
            return;

        if (value.TryGetText(out string text))
            loadBudget.AddCellText(text);
        if (formula is not null)
            loadBudget.AddCellText(formula);
    }

    public void OnAutoFilter(RangeRef range, bool hasCriteria)
    {
        sheet.AutoFilterRange = range;
        sheet.AutoFilterIsActive = hasCriteria;
    }

    public void OnDataValidation(DataValidation validation) => sheet.Validations.AddFromFile(validation);

    public void OnConditionalFormat(ConditionalFormatRule rule) => sheet.ConditionalFormats.AddFromFile(rule);

    public void OnCellMarkup(CellRef cell, List<KeyValuePair<string, string>> attributes) =>
        sheet.CellMarkupTable[cell] = attributes;

    public void OnFormulaMarkup(CellRef cell, FormulaMarkup markup)
    {
        sheet.FormulaMarkupTable[cell] = markup;
        if (!markup.WritesOwnElement || _warnedAboutDataTables)
            return;

        // The element travels intact, but nothing recalculates it: the engine has no model of a
        // What-If table, so its cached values go stale the moment an input changes.
        _warnedAboutDataTables = true;
        warn(new WorkbookWarning(
            "DataTableFormula",
            $"'{sheet.Name}' holds a What-If data table. It is written back unchanged but never recalculated.",
            cell));
    }

    public PageSetup PageSetup => sheet.PageSetup;

    public SheetProtection Protection => sheet.Protection;

    public void OnMerge(RangeRef range)
    {
        try
        {
            sheet.Merge(range);
        }
        catch (InvalidOperationException)
        {
            warn(new WorkbookWarning("OverlappingMerge", $"Merged range {range} overlaps another merge and was dropped."));
        }
    }

    public void OnFormulaGroup(FormulaGroup group, int? sharedIndex)
    {
        sheet.FormulaGroups.Add(group);
        if (sharedIndex is { } index)
            _sharedGroups[index] = group;
    }

    public void OnSharedFormulaMember(CellRef cell, int sharedIndex)
    {
        // A master declares the extent of its group, but some producers omit or understate the
        // ref attribute, so members outside it widen the range.
        if (!_sharedGroups.TryGetValue(sharedIndex, out FormulaGroup? group))
            return;

        group.Members.Add(cell);
        if (!group.Range.Contains(cell))
            group.Range = new RangeRef(group.Range.Start, cell);
    }

    /// <summary>Whether a shared formula group with this index has been seen.</summary>
    public bool TryGetSharedGroup(int sharedIndex, out FormulaGroup? group) =>
        _sharedGroups.TryGetValue(sharedIndex, out group);
}
