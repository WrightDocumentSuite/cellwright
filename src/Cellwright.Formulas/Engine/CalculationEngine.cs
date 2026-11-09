using System.Diagnostics;
using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas.Engine;

/// <summary>
/// Drives recalculation of a workbook: keeps the dependency graph, tracks which formulas went
/// stale and evaluates them in dependency order. One engine is attached per workbook and
/// survives between calls so repeated edits only recompute what they touched.
/// </summary>
internal sealed class CalculationEngine
{
    private readonly Workbook _workbook;
    private readonly ProgramCache _programs;
    private readonly EvalContext _context;
    private readonly Evaluator _evaluator = new();
    private readonly HashSet<int> _dirty = [];

    private DependencyIndex? _index;
    private int[] _order = [];
    private HashSet<int> _circular = [];
    private bool _structureChanged = true;
    private bool _firstPass = true;

    public CalculationEngine(Workbook workbook)
    {
        _workbook = workbook;
        _programs = new ProgramCache(workbook);
        _context = new EvalContext(workbook, static (sheet, row, column) => sheet.Values.Get(row, column));
        workbook.CellChanged += OnCellChanged;
        workbook.VisibilityChanged += OnVisibilityChanged;
        workbook.StructureChanged += OnStructureChanged;
        workbook.Names.Changed += OnNamesChanged;
    }

    /// <summary>Materialized ranges retained between passes; inspected by tests and benchmarks.</summary>
    internal RangeCache Ranges => _context.Ranges;

    /// <summary>Number of distinct formula texts retained by all compiled-program caches.</summary>
    internal int ProgramCount => _programs.Count + _context.NestedProgramCount;

    /// <summary>Recalculates the workbook and reports what happened.</summary>
    public CalculationSummary Calculate(CalculateOptions options)
    {
        long start = Stopwatch.GetTimestamp();

        // Stays false unless the pass finishes, so a cancelled or failed one is not mistaken for
        // a complete result by CalculatedValue or by fullCalcOnLoad on save.
        _workbook.IsFullyCalculated = false;
        _context.BeginPass();
        _context.CancellationToken = options.CancellationToken;
        _context.OnWarning = options.OnWarning;
        _context.FormatCulture = options.FormatCulture ?? System.Globalization.CultureInfo.InvariantCulture;

        SharedFormulaExpander.Expand(_workbook, _programs, options.OnWarning);
        EnsureIndex();

        int calculated = 0;
        int errors = 0;
        int[] targets = SelectTargetsInOrder(options.Force);

        // A spill writes cells no edge of the graph predicted — the rectangle exists only once
        // the formula has run — so anything reading those cells may have gone before them. The
        // pass is repeated over just the readers of what changed. A round is only caused by a
        // spill that changed in the round before, and the bound stops arrays that feed themselves.
        for (int pass = 0; pass < MaxSpillPasses && targets.Length > 0; pass++)
        {
            _spilledCells.Clear();
            foreach (int entryIndex in targets)
            {
                options.CancellationToken.ThrowIfCancellationRequested();
                calculated++;
                if (Evaluate(_index!.Entries[entryIndex], _circular.Contains(entryIndex)).IsError)
                    errors++;
            }

            targets = SpillReaders();
        }

        // Only now has the work been done: until this point a retry must redo all of it.
        _firstPass = false;
        _dirty.Clear();
        ReportCircular(options.OnWarning);

        bool settled = targets.Length == 0;
        if (!settled)
        {
            options.OnWarning?.Invoke(new WorkbookWarning(
                "SpillDidNotSettle",
                $"Dynamic arrays were still changing after {MaxSpillPasses} rounds; some formulas may hold outdated values."));
        }

        _workbook.IsFullyCalculated = _circular.Count == 0 && settled;

        return new CalculationSummary(calculated, errors, DescribeCircular(), Stopwatch.GetElapsedTime(start));
    }

    /// <summary>Evaluates a formula outside any cell, as <c>Workbook.Evaluate</c> does.</summary>
    public CellValue EvaluateAdHoc(string formula, Worksheet sheet, CellRef cell)
    {
        FormulaProgram? program = _programs.GetOrParse(formula, out FormulaParseException? failure);
        if (program is null)
            throw failure!;

        BeginAdHocPass();
        _context.MoveTo(sheet, cell);
        FormulaValue value = _evaluator.Evaluate(program, _context);
        return Operators.Dereference(value, _context);
    }

    /// <summary>
    /// Evaluates one stored formula as copied from <paramref name="source"/> to
    /// <paramref name="target"/>. The source text is the cache key, while relative references
    /// receive the copy offset only during execution.
    /// </summary>
    public CellValue EvaluateAdHocCopied(
        string formula, Worksheet sheet, CellRef source, CellRef target)
    {
        FormulaProgram? program = _programs.GetOrParse(formula, out FormulaParseException? failure);
        if (program is null)
            throw failure!;

        BeginAdHocPass();
        _context.MoveTo(sheet, target);
        FormulaValue value = _evaluator.Evaluate(
            program,
            _context,
            target.Row - source.Row,
            target.Column - source.Column);
        return Operators.Dereference(value, _context);
    }

    /// <summary>
    /// Starts an independent one-formula calculation. A preceding workbook calculation may have
    /// carried cancellation, warning and authoring-culture options; public ad-hoc evaluation must
    /// not inherit that unrelated request's state.
    /// </summary>
    private void BeginAdHocPass()
    {
        _context.BeginPass();
        _context.CancellationToken = CancellationToken.None;
        _context.OnWarning = null;
        _context.FormatCulture = System.Globalization.CultureInfo.InvariantCulture;
    }

    /// <summary>How many rounds a pass may repeat over the readers of moved spills.</summary>
    private const int MaxSpillPasses = 32;

    /// <summary>The cells every spill of this pass wrote or vacated, for the follow-up round.</summary>
    private readonly List<(Worksheet Sheet, RangeRef Range)> _spilledCells = [];

    private CellValue Evaluate(FormulaEntry entry, bool isCircular)
    {
        CellValue result;
        if (isCircular)
        {
            // Excel leaves circular cells at zero rather than failing the whole workbook.
            result = CellValue.Number(0);
        }
        else
        {
            _context.MoveTo(entry.Sheet, entry.Cell);
            FormulaValue value = _evaluator.Evaluate(entry.Program, _context);

            // An array of more than one value spills, unless the formula is the master of a
            // legacy CSE group — those keep their pre-2018 meaning, one cell however large
            // the array.
            if (value.IsArray && value.AsArray is { Count: > 1 } array && !IsLegacyArrayMaster(entry))
                return Spill(entry, array);

            result = Operators.Dereference(value, _context);

            // A formula never leaves a cell truly empty: pointing at a blank cell shows zero,
            // which is also what Excel caches in the file.
            if (result.IsBlank)
                result = CellValue.Number(0);
        }

        RetractSpill(entry);
        entry.Sheet.Cell(entry.Cell).SetComputedValue(result);

        // Anything that already read a range covering this cell now holds a stale copy.
        _context.Ranges.Invalidate(entry.Sheet, entry.Cell.Row, entry.Cell.Column);
        return result;
    }

    /// <summary>
    /// Publishes an array over its rectangle: the anchor takes the first value, its neighbours
    /// take the rest, and the sheet records the rectangle. A cell in the way makes the whole
    /// formula <c>#SPILL!</c> rather than a partial fill — nothing of the array is shown, which
    /// is Excel's behaviour and the honest one, since a truncated array reads as the answer.
    /// </summary>
    private CellValue Spill(FormulaEntry entry, ArrayValue array)
    {
        Worksheet sheet = entry.Sheet;
        CellRef anchor = entry.Cell;

        int lastRow = anchor.Row + array.Rows - 1;
        int lastColumn = anchor.Column + array.Columns - 1;
        RangeRef? previous = sheet.SpillTable.TryGetValue(anchor, out RangeRef held) ? held : null;

        if (lastRow > SheetLimits.MaxRow || lastColumn > SheetLimits.MaxColumn ||
            IsBlocked(sheet, anchor, lastRow, lastColumn, previous))
        {
            RetractSpill(entry);
            CellValue error = CellValue.Error(ExcelError.Spill);
            sheet.Cell(anchor).SetComputedValue(error);
            _context.Ranges.Invalidate(sheet, anchor.Row, anchor.Column);
            return error;
        }

        var target = new RangeRef(anchor, new CellRef(lastRow, lastColumn));

        // Cells of the old rectangle that the new one no longer covers go back to blank.
        if (previous is { } old)
            VacateOutside(sheet, old, target);

        // No edge of the graph leads from the anchor to a reader of one of its spilled cells, so
        // new values in an unchanged rectangle need the follow-up round as much as a moved one.
        bool changed = previous != target;
        for (int row = 0; row < array.Rows; row++)
        {
            for (int column = 0; column < array.Columns; column++)
            {
                CellValue value = array[row, column];
                if (value.IsBlank)
                    value = CellValue.Number(0);

                changed |= sheet.Values.Get(anchor.Row + row, anchor.Column + column) != value;
                sheet.Values.Set(anchor.Row + row, anchor.Column + column, value);
                _context.Ranges.Invalidate(sheet, anchor.Row + row, anchor.Column + column);
            }
        }

        sheet.SpillTable[anchor] = target;
        if (changed)
        {
            _spilledCells.Add((sheet, target));
            if (previous is { } vacated && vacated != target)
                _spilledCells.Add((sheet, vacated));
        }

        return array[0, 0];
    }

    /// <summary>Whether anything stands where the array wants to land.</summary>
    private static bool IsBlocked(Worksheet sheet, CellRef anchor, int lastRow, int lastColumn, RangeRef? own)
    {
        for (int row = anchor.Row; row <= lastRow; row++)
        {
            for (int column = anchor.Column; column <= lastColumn; column++)
            {
                var cell = new CellRef(row, column);
                if (cell == anchor)
                    continue;

                if (sheet.Formulas.Get(row, column) is not null)
                    return true;

                // A value is only in the way when it is somebody's: a cell this anchor spilled
                // into last time holds this formula's own output and may be written over.
                if (!sheet.Values.Get(row, column).IsBlank && (own is not { } held || !held.Contains(cell)))
                    return true;

                if (sheet.SpillAnchorCovering(cell) is { } owner && owner != anchor)
                    return true;
            }
        }

        // A merge anywhere in the rectangle blocks it whole, as it does in Excel.
        foreach (RangeRef merge in sheet.MergedRanges)
        {
            if (merge.Start.Row <= lastRow && merge.End.Row >= anchor.Row &&
                merge.Start.Column <= lastColumn && merge.End.Column >= anchor.Column)
                return true;
        }

        return false;
    }

    /// <summary>Takes back a formula's spill, blanking every cell of it except the anchor.</summary>
    private void RetractSpill(FormulaEntry entry)
    {
        if (!entry.Sheet.SpillTable.TryGetValue(entry.Cell, out RangeRef range))
            return;

        entry.Sheet.SpillTable.Remove(entry.Cell);
        VacateOutside(entry.Sheet, range, new RangeRef(entry.Cell, entry.Cell));
        _spilledCells.Add((entry.Sheet, range));
    }

    /// <summary>Blanks the cells of a rectangle that another rectangle does not cover.</summary>
    private void VacateOutside(Worksheet sheet, RangeRef old, RangeRef kept)
    {
        for (int row = old.Start.Row; row <= old.End.Row; row++)
        {
            for (int column = old.Start.Column; column <= old.End.Column; column++)
            {
                if (kept.Contains(new CellRef(row, column)))
                    continue;

                sheet.Values.Set(row, column, CellValue.Blank);
                _context.Ranges.Invalidate(sheet, row, column);
            }
        }
    }

    /// <summary>Whether a formula is the master of a legacy CSE array group.</summary>
    private static bool IsLegacyArrayMaster(FormulaEntry entry)
    {
        foreach (FormulaGroup group in entry.Sheet.FormulaGroups)
        {
            if (group.Kind == FormulaGroupKind.Array && group.Range.Contains(entry.Cell))
                return true;
        }

        return false;
    }

    /// <summary>
    /// The formulas that read any cell a spill just wrote or vacated, in dependency order —
    /// the targets of the follow-up round.
    /// </summary>
    private int[] SpillReaders()
    {
        if (_spilledCells.Count == 0)
            return [];

        var readers = new HashSet<int>();
        foreach ((Worksheet sheet, RangeRef range) in _spilledCells)
        {
            for (int row = range.Start.Row; row <= range.End.Row; row++)
            {
                for (int column = range.Start.Column; column <= range.End.Column; column++)
                    _index!.CollectDependents(sheet, new CellRef(row, column), readers);
            }
        }

        if (readers.Count == 0)
            return [];

        _index!.ExpandTransitively(readers);
        int[] ordered = [.. readers];
        int[] positions = Array.ConvertAll(ordered, _index.PositionOf);
        Array.Sort(positions, ordered);
        return ordered;
    }

    /// <summary>
    /// The formulas to evaluate, in dependency order: everything on the first pass or when
    /// forced, otherwise the cells touched since last time, the volatile ones and whatever
    /// depends on them. The partial case never scans the whole workbook.
    /// </summary>
    private int[] SelectTargetsInOrder(bool force)
    {
        if (force || _firstPass)
            return _order;

        var targets = new HashSet<int>(_dirty);
        foreach (int entry in _index!.VolatileEntries)
            targets.Add(entry);
        if (targets.Count == 0)
            return [];

        _index.ExpandTransitively(targets);
        int[] ordered = [.. targets];
        int[] positions = Array.ConvertAll(ordered, _index.PositionOf);
        Array.Sort(positions, ordered);
        return ordered;
    }

    private void EnsureIndex()
    {
        if (_index is not null && !_structureChanged)
            return;

        _index = new DependencyIndex();
        _context.Ranges.Clear();
        foreach (Worksheet sheet in _workbook.Sheets)
        {
            RetractOrphanedSpills(sheet);
            AddSheetFormulas(sheet);
        }

        _index.Connect(_context);
        (_order, _circular) = CalculationOrder.Build(_index.Entries);
        _index.SetOrder(_order);
        _structureChanged = false;
        _firstPass = true;
    }

    /// <summary>
    /// A spill belongs to the formula at its anchor. When that formula has been replaced by a
    /// value or cleared, nothing will ever recalculate the rectangle, so it is taken back here.
    /// </summary>
    private void RetractOrphanedSpills(Worksheet sheet)
    {
        if (sheet.SpillTable.Count == 0)
            return;

        foreach ((CellRef anchor, RangeRef range) in sheet.SpillTable.ToArray())
        {
            if (sheet.Formulas.Get(anchor.Row, anchor.Column) is not null)
                continue;

            sheet.SpillTable.Remove(anchor);
            VacateOutside(sheet, range, new RangeRef(anchor, anchor));
        }
    }

    private void AddSheetFormulas(Worksheet sheet)
    {
        foreach ((int row, int column, string? text) in sheet.Formulas.EnumerateOccupied())
        {
            if (text is null)
                continue;

            FormulaProgram? program = _programs.GetOrParse(text, out FormulaParseException? failure);
            var cell = new CellRef(row, column);
            if (program is null)
            {
                _context.OnWarning?.Invoke(new WorkbookWarning("BadFormula", failure!.Message, cell));
                sheet.Cell(cell).SetComputedValue(CellValue.Error(ExcelError.Name));
                continue;
            }

            _index!.AddFormula(sheet, cell, program);
        }
    }

    private void OnCellChanged(Worksheet sheet, CellRef cell, bool structural)
    {
        // An edit through the model invalidates cached rectangles exactly like a computed
        // result does; a structural change drops the whole cache when the index is rebuilt.
        _context.Ranges.Invalidate(sheet, cell.Row, cell.Column);

        // Typing into a spilled cell breaks the spill: the rectangle is forgotten, so the next
        // recalculation of its anchor finds a foreign value in the way and answers #SPILL!.
        // Only user edits arrive here — the engine's own writes never raise the event — which
        // is what tells somebody's value apart from the formula's own output.
        if (sheet.SpillAnchorCovering(cell) is { } anchor && anchor != cell)
        {
            sheet.SpillTable.Remove(anchor);
            if (_index?.IndexOf(sheet, anchor.Row, anchor.Column) is { } entry and >= 0)
                _dirty.Add(entry);
        }

        if (structural)
        {
            _structureChanged = true;
            return;
        }

        if (_index is null)
            return;

        _index.CollectDependents(sheet, cell, _dirty);
    }

    /// <summary>
    /// Hiding a row writes no cell, so nothing lands in the dirty set through the usual path.
    /// Only SUBTOTAL and AGGREGATE can notice, and the index keeps a list of exactly those.
    /// </summary>
    private void OnVisibilityChanged(Worksheet sheet)
    {
        if (_index is null)
            return;

        foreach (int entry in _index.VisibilityDependentEntries)
            _dirty.Add(entry);
    }

    /// <summary>Sheet references are bound when a formula compiles, so the compiled forms go too.</summary>
    private void OnStructureChanged()
    {
        _programs.Clear();
        _context.ClearPrograms();
        _structureChanged = true;
    }

    private void OnNamesChanged()
    {
        _workbook.IsFullyCalculated = false;
        _structureChanged = true;
    }

    private void ReportCircular(Action<WorkbookWarning>? onWarning)
    {
        if (onWarning is null || _circular.Count == 0)
            return;

        foreach (int entryIndex in _circular)
        {
            FormulaEntry entry = _index!.Entries[entryIndex];
            onWarning(new WorkbookWarning(
                "CircularReference",
                $"'{entry.Sheet.Name}'!{entry.Cell} takes part in a circular reference and was left at zero.",
                entry.Cell));
        }
    }

    private IReadOnlyList<string> DescribeCircular()
    {
        if (_circular.Count == 0)
            return [];

        var names = new List<string>(_circular.Count);
        foreach (int entryIndex in _circular)
        {
            FormulaEntry entry = _index!.Entries[entryIndex];
            names.Add($"{entry.Sheet.Name}!{entry.Cell}");
        }

        names.Sort(StringComparer.Ordinal);
        return names;
    }
}
