using Cellwright.Formulas.Functions;
using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas.Engine;

/// <summary>Identity of a cell across the workbook.</summary>
internal readonly record struct CellKey(int SheetId, int Row, int Column);

/// <summary>One formula cell tracked by the engine.</summary>
internal sealed class FormulaEntry(Worksheet sheet, CellRef cell, FormulaProgram program)
{
    public Worksheet Sheet { get; } = sheet;

    public CellRef Cell { get; } = cell;

    public FormulaProgram Program { get; } = program;

    /// <summary>Indexes of the formula cells this one reads.</summary>
    public List<int> Precedents { get; } = [];
}

/// <summary>
/// The dependency graph of a workbook's formulas. Built in one pass after loading, because
/// walking every formula once is far cheaper than maintaining the graph cell by cell while a
/// file is being read.
/// </summary>
internal sealed class DependencyIndex
{
    private readonly List<FormulaEntry> _entries = [];
    private readonly Dictionary<CellKey, int> _byCell = [];
    private readonly Dictionary<int, Dictionary<int, List<int>>> _rowsByColumn = [];
    private readonly Dictionary<CellKey, List<int>> _pointDependents = [];
    private readonly RangeSubscriptionIndex.Builder _subscriptions = new();
    private readonly List<int> _volatile = [];
    private readonly List<int> _visibilityDependent = [];
    private RangeSubscriptionIndex? _rangeDependents;
    private List<int>?[] _dependents = [];
    private int[] _positions = [];

    /// <summary>
    /// Functions that look at where a reference points rather than at what it holds. A formula
    /// using one can name its own cell without being circular: <c>ROW(A1)</c> in A1 is ordinary,
    /// and so is <c>INDEX(A:A,ROW()-1)</c> at the bottom of a column.
    /// </summary>
    private static readonly HashSet<string> AddressOnlyFunctions = new(StringComparer.Ordinal)
    {
        "ROW", "COLUMN", "ROWS", "COLUMNS", "AREAS", "ISREF", "ISFORMULA", "FORMULATEXT",
        "INDEX", "OFFSET", "CELL", "SHEET", "SHEETS",
    };

    /// <summary>Every formula cell, in discovery order.</summary>
    public IReadOnlyList<FormulaEntry> Entries => _entries;

    /// <summary>Adds a formula cell before the graph is wired up.</summary>
    public void AddFormula(Worksheet sheet, CellRef cell, FormulaProgram program)
    {
        var key = new CellKey(sheet.Id, cell.Row, cell.Column);
        if (_byCell.ContainsKey(key))
            return;

        _byCell[key] = _entries.Count;
        _entries.Add(new FormulaEntry(sheet, cell, program));

        Dictionary<int, List<int>> columns = _rowsByColumn.TryGetValue(sheet.Id, out Dictionary<int, List<int>>? existing)
            ? existing
            : _rowsByColumn[sheet.Id] = [];
        if (!columns.TryGetValue(cell.Column, out List<int>? rows))
            columns[cell.Column] = rows = [];
        rows.Add(cell.Row);
    }

    /// <summary>Looks up the entry index of a formula cell, or -1.</summary>
    public int IndexOf(Worksheet sheet, int row, int column) =>
        _byCell.TryGetValue(new CellKey(sheet.Id, row, column), out int index) ? index : -1;

    /// <summary>
    /// Resolves every reference of every formula into graph edges. Ranges only produce edges
    /// for the formula cells they actually contain, found through a per-column index.
    /// </summary>
    public void Connect(EvalContext context)
    {
        foreach (Dictionary<int, List<int>> columns in _rowsByColumn.Values)
        {
            foreach (List<int> rows in columns.Values)
                rows.Sort();
        }

        var visitedNames = new HashSet<DefinedName>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < _entries.Count; i++)
        {
            FormulaEntry entry = _entries[i];
            context.MoveTo(entry.Sheet, entry.Cell);
            visitedNames.Clear();
            ConnectProgram(i, entry, entry.Program, context, visitedNames, nameDepth: 0);
        }

        _rangeDependents = _subscriptions.Build();
        BuildReverseEdges();
    }

    /// <summary>
    /// Wires one program into the graph. A defined name is part of the formula that uses it, so
    /// its references, volatility and nested names count as the formula's own; lookup follows
    /// the scope of the using cell, as evaluation does.
    /// </summary>
    private void ConnectProgram(
        int entryIndex,
        FormulaEntry entry,
        FormulaProgram program,
        EvalContext context,
        HashSet<DefinedName> visitedNames,
        int nameDepth)
    {
        foreach (FormulaRef reference in program.References)
        {
            // A spill reference names a rectangle that exists only after its anchor has
            // run, so the edge goes to the anchor cell: when the anchor recalculates and
            // its rectangle moves, the follow-up round reaches this formula through it.
            if (reference.IsSpilled)
            {
                Connect(entryIndex, entry, program, ResolveSpillAnchor(reference, entry.Sheet, context));
                continue;
            }

            Connect(entryIndex, entry, program, context.Resolve(reference));
        }

        // A structured reference resolves against the table's current shape. That shape only
        // changes structurally, and a structural change rebuilds this index anyway, so the
        // area it names now is the area it will name until then.
        foreach (TableRef reference in program.TableReferences)
            Connect(entryIndex, entry, program, context.Resolve(reference));

        // Entries are wired one at a time, so a repeat for this entry can only be the last item.
        if (program.IsVolatile && (_volatile.Count == 0 || _volatile[^1] != entryIndex))
            _volatile.Add(entryIndex);
        if (program.DependsOnVisibility && (_visibilityDependent.Count == 0 || _visibilityDependent[^1] != entryIndex))
            _visibilityDependent.Add(entryIndex);

        if (nameDepth >= NameResolver.MaxNameDepth)
            return;

        foreach (Instruction instruction in program.Instructions)
        {
            if (instruction.Code != OpCode.PushName)
                continue;

            DefinedName? defined = context.Workbook.Names.Resolve(program.Strings[instruction.Operand], context.Sheet);
            if (defined is not null && visitedNames.Add(defined) && context.GetNameProgram(defined.RefersTo) is { } named)
                ConnectProgram(entryIndex, entry, named, context, visitedNames, nameDepth + 1);
        }
    }

    private static bool OnlyInspectsAddresses(FormulaProgram program)
    {
        foreach (FunctionDefinition? function in program.Functions)
        {
            if (function is not null && AddressOnlyFunctions.Contains(function.Name))
                return true;
        }

        return false;
    }

    private void BuildReverseEdges()
    {
        _dependents = new List<int>?[_entries.Count];
        var seen = new HashSet<int>();
        for (int i = 0; i < _entries.Count; i++)
        {
            // A formula like SUM(A1:A5)+A3 lists A3 twice; duplicates would double the reverse
            // edges and the work of every dirty-set expansion.
            List<int> precedents = _entries[i].Precedents;
            seen.Clear();
            for (int p = precedents.Count - 1; p >= 0; p--)
            {
                if (!seen.Add(precedents[p]))
                    precedents.RemoveAt(p);
            }

            foreach (int precedent in precedents)
                (_dependents[precedent] ??= []).Add(i);
        }
    }

    /// <summary>The anchor cell of a spill reference, as the one-cell area the edge points at.</summary>
    private static ReferenceValue? ResolveSpillAnchor(in FormulaRef reference, Worksheet own, EvalContext context)
    {
        Worksheet? sheet = reference.SheetId switch
        {
            FormulaRef.CurrentSheet => own,
            FormulaRef.UnresolvedSheet => null,
            _ => context.Workbook.Sheets.FindById(reference.SheetId),
        };

        return sheet is null
            ? null
            : new ReferenceValue(
                sheet, reference.Start.Row, reference.Start.Column, reference.Start.Row, reference.Start.Column);
    }

    private void Connect(int entryIndex, FormulaEntry entry, FormulaProgram program, ReferenceValue? reference)
    {
        // A 3-D reference resolves to a chain of areas, one per spanned sheet; each of them
        // subscribes on its own.
        for (ReferenceValue? area = reference; area is not null; area = area.Next)
            ConnectArea(entryIndex, entry, program, area);
    }

    private void ConnectArea(int entryIndex, FormulaEntry entry, FormulaProgram program, ReferenceValue area)
    {
        if (area.FirstRow == area.LastRow && area.FirstColumn == area.LastColumn)
        {
            var key = new CellKey(area.Sheet.Id, area.FirstRow, area.FirstColumn);
            AddDependent(_pointDependents, key, entryIndex);

            // A self-edge is kept on purpose: it is exactly how a cell like =A1+1 in A1
            // announces its own circular reference to the ordering pass.
            int precedent = IndexOf(area.Sheet, area.FirstRow, area.FirstColumn);
            if (precedent >= 0 && (precedent != entryIndex || !OnlyInspectsAddresses(program)))
                entry.Precedents.Add(precedent);
            return;
        }

        _subscriptions.Add(
            new RangeKey(area.Sheet.Id, area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn),
            entryIndex);

        if (!_rowsByColumn.TryGetValue(area.Sheet.Id, out Dictionary<int, List<int>>? columns))
            return;

        for (int column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            if (!columns.TryGetValue(column, out List<int>? rows))
                continue;

            int start = rows.BinarySearch(area.FirstRow);
            if (start < 0)
                start = ~start;

            for (int r = start; r < rows.Count && rows[r] <= area.LastRow; r++)
            {
                int precedent = IndexOf(area.Sheet, rows[r], column);
                if (precedent >= 0 && (precedent != entryIndex || !OnlyInspectsAddresses(program)))
                    entry.Precedents.Add(precedent);
            }
        }
    }

    /// <summary>Collects the formulas that must be recalculated after a cell changes.</summary>
    public void CollectDependents(Worksheet sheet, CellRef cell, HashSet<int> destination)
    {
        var key = new CellKey(sheet.Id, cell.Row, cell.Column);
        if (_pointDependents.TryGetValue(key, out List<int>? direct))
        {
            foreach (int entry in direct)
                destination.Add(entry);
        }

        _rangeDependents?.Collect(sheet.Id, cell.Row, cell.Column, destination);
    }

    /// <summary>
    /// Adds every entry that transitively depends on the ones already in the set. Walking the
    /// reverse edges keeps the cost proportional to the affected subgraph rather than to the
    /// size of the workbook, which is the whole point of recalculating incrementally.
    /// </summary>
    public void ExpandTransitively(HashSet<int> dirty)
    {
        var queue = new Queue<int>(dirty);
        while (queue.Count > 0)
        {
            int entry = queue.Dequeue();
            if (_dependents[entry] is not { } dependents)
                continue;

            foreach (int dependent in dependents)
            {
                if (dirty.Add(dependent))
                    queue.Enqueue(dependent);
            }
        }
    }

    /// <summary>Entries whose formulas must be recomputed on every pass.</summary>
    public IReadOnlyList<int> VolatileEntries => _volatile;

    /// <summary>Entries whose result depends on which rows are hidden.</summary>
    public IReadOnlyList<int> VisibilityDependentEntries => _visibilityDependent;

    /// <summary>Position of an entry in the calculation order, for cheap ordered iteration.</summary>
    public int PositionOf(int entry) => _positions[entry];

    /// <summary>Records the calculation order so dirty subsets can be walked without a full scan.</summary>
    public void SetOrder(int[] order)
    {
        _positions = new int[_entries.Count];
        for (int position = 0; position < order.Length; position++)
            _positions[order[position]] = position;
    }

    private static void AddDependent(Dictionary<CellKey, List<int>> map, CellKey key, int entry)
    {
        if (!map.TryGetValue(key, out List<int>? list))
            map[key] = list = [];
        list.Add(entry);
    }
}
