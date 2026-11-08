using Cellwright.Formulas.Parsing;

namespace Cellwright.Formulas;

/// <summary>
/// Everything a function needs to know about where it is running: the workbook, the cell that
/// owns the formula and the services for reading other cells.
/// </summary>
internal sealed class EvalContext
{
    private readonly Func<Worksheet, int, int, CellValue> _readCell;
    private readonly Dictionary<string, FormulaProgram?> _namePrograms = new(StringComparer.Ordinal);
    private readonly Dictionary<string, LinkedListNode<DynamicProgram>> _dynamicPrograms =
        new(StringComparer.Ordinal);
    private readonly LinkedList<DynamicProgram> _dynamicLru = [];
    private readonly Stack<Evaluator> _spareEvaluators = new();

    /// <summary>
    /// Maximum distinct formulas retained for dynamic text such as INDIRECT targets. A copied
    /// formula can generate one target per worksheet row, so this cache must be bounded.
    /// </summary>
    internal const int MaxDynamicPrograms = 1_024;

    public EvalContext(Workbook workbook, Func<Worksheet, int, int, CellValue> readCell)
    {
        Workbook = workbook;
        _readCell = readCell;
        Timestamp = DateTime.Now;
    }

    /// <summary>The workbook being calculated.</summary>
    public Workbook Workbook { get; }

    /// <summary>Sheet holding the formula under evaluation.</summary>
    public Worksheet Sheet { get; private set; } = null!;

    /// <summary>Cell holding the formula under evaluation.</summary>
    public CellRef Cell { get; private set; }

    /// <summary>Date system of the workbook.</summary>
    public ExcelDateSystem DateSystem => Workbook.DateSystem;

    /// <summary>
    /// The moment the calculation pass started. Frozen so that <c>NOW</c> is consistent across
    /// every formula in one pass, the way Excel behaves.
    /// </summary>
    public DateTime Timestamp { get; private set; }

    /// <summary>Cancellation for long recalculations.</summary>
    public CancellationToken CancellationToken { get; set; }

    /// <summary>Warnings raised while evaluating, e.g. clamped ranges.</summary>
    public Action<WorkbookWarning>? OnWarning { get; set; }

    /// <summary>
    /// Culture whose separators <c>TEXT</c> format strings are written in — the authoring
    /// locale of the file, invariant unless a calculation pass says otherwise.
    /// </summary>
    public System.Globalization.CultureInfo FormatCulture { get; set; } = System.Globalization.CultureInfo.InvariantCulture;

    /// <summary>
    /// Ranges already read out of the worksheets during this session. Lives as long as the
    /// engine: entries survive between passes and are dropped the moment something is written
    /// inside them.
    /// </summary>
    public RangeCache Ranges { get; } = new();

    /// <summary>How deep defined-name resolution currently is; bounded to stop name cycles.</summary>
    public int NameDepth { get; set; }

    /// <summary>How deep <c>LAMBDA</c> invocation currently is; bounded to stop infinite recursion.</summary>
    public int LambdaDepth { get; set; }

    /// <summary>
    /// Compiles the formula of a defined name, reusing earlier results. This cache is naturally
    /// bounded by the workbook's finite names collection.
    /// </summary>
    internal FormulaProgram? GetNameProgram(string text)
    {
        if (_namePrograms.TryGetValue(text, out FormulaProgram? cached))
            return cached;

        FormulaProgram? program = ParseProgram(text);
        _namePrograms[text] = program;
        return program;
    }

    /// <summary>
    /// Compiles formula text produced dynamically by INDIRECT. Both successful and failed parses
    /// live in an exact LRU so a high-cardinality formula cannot retain a program per cell while
    /// the common repeated-target case still parses once.
    /// </summary>
    internal FormulaProgram? GetDynamicProgram(string text)
    {
        if (_dynamicPrograms.TryGetValue(text, out LinkedListNode<DynamicProgram>? cached))
        {
            _dynamicLru.Remove(cached);
            _dynamicLru.AddLast(cached);
            return cached.Value.Program;
        }

        var entry = new DynamicProgram(text, ParseProgram(text));
        LinkedListNode<DynamicProgram> node = _dynamicLru.AddLast(entry);
        _dynamicPrograms.Add(text, node);
        if (_dynamicPrograms.Count > MaxDynamicPrograms)
        {
            LinkedListNode<DynamicProgram> oldest = _dynamicLru.First!;
            _dynamicLru.RemoveFirst();
            _dynamicPrograms.Remove(oldest.Value.Text);
        }

        return entry.Program;
    }

    /// <summary>Programs retained outside the engine's primary formula cache.</summary>
    internal int NestedProgramCount => _namePrograms.Count + _dynamicPrograms.Count;

    private FormulaProgram? ParseProgram(string text)
    {
        try
        {
            return FormulaParser.Parse(text, Workbook);
        }
        catch (FormulaParseException)
        {
            return null;
        }
    }

    private sealed record DynamicProgram(string Text, FormulaProgram? Program);

    /// <summary>
    /// Runs a program inside the formula currently being evaluated. Evaluators own a reusable
    /// stack, so they are pooled rather than allocated per nested call.
    /// </summary>
    internal FormulaValue EvaluateNested(FormulaProgram program)
    {
        Evaluator evaluator = _spareEvaluators.Count > 0 ? _spareEvaluators.Pop() : new Evaluator();
        try
        {
            return evaluator.Evaluate(program, this);
        }
        finally
        {
            _spareEvaluators.Push(evaluator);
        }
    }

    internal FormulaValue InvokeLambda(LambdaValue lambda, ReadOnlySpan<FormulaValue> args)
    {
        const int MaxLambdaDepth = 64;
        if (LambdaDepth >= MaxLambdaDepth)
            return FormulaValue.Error(ExcelError.Number);

        int[] slots = lambda.Definition.ParameterSlots;
        if (args.Length > slots.Length || (args.Length == 0 && slots.Length > 0))
            return FormulaValue.Error(ExcelError.Value);

        LambdaDepth++;
        Evaluator evaluator = _spareEvaluators.Count > 0 ? _spareEvaluators.Pop() : new Evaluator();
        try
        {
            return evaluator.EvaluateLambda(lambda, args, this);
        }
        finally
        {
            _spareEvaluators.Push(evaluator);
            LambdaDepth--;
        }
    }

    /// <summary>Points the context at the next formula to evaluate.</summary>
    public void MoveTo(Worksheet sheet, CellRef cell)
    {
        Sheet = sheet;
        Cell = cell;
    }

    /// <summary>Restarts the frozen timestamp at the beginning of a calculation pass.</summary>
    public void BeginPass() => Timestamp = DateTime.Now;

    /// <summary>Reads a cell's current value.</summary>
    public CellValue ReadCell(Worksheet sheet, int row, int column) => _readCell(sheet, row, column);

    /// <summary>Reports a recoverable problem.</summary>
    public void Warn(string code, string message) => OnWarning?.Invoke(new WorkbookWarning(code, message, Cell));

    /// <summary>
    /// Turns a compiled reference into a concrete area. Whole-column and whole-row references
    /// are clamped to the used range so that <c>SUM(A:A)</c> does not walk a million rows.
    /// A 3-D span becomes a chain of areas, one per sheet between its endpoints in tab order.
    /// </summary>
    public ReferenceValue? Resolve(in FormulaRef reference)
    {
        Worksheet? sheet = reference.SheetId switch
        {
            FormulaRef.CurrentSheet => Sheet,
            FormulaRef.UnresolvedSheet => null,
            _ => Workbook.Sheets.FindById(reference.SheetId),
        };

        if (sheet is null)
            return null;

        // A1# names whatever the dynamic array at A1 spilled into. A cell that spilled
        // nothing has no rectangle to name, which reads as #REF! — Excel's answer too.
        if (reference.IsSpilled)
        {
            var anchor = new CellRef(reference.Start.Row, reference.Start.Column);
            return sheet.SpillTable.TryGetValue(anchor, out RangeRef spill)
                ? new ReferenceValue(sheet, spill.Start.Row, spill.Start.Column, spill.End.Row, spill.End.Column)
                : null;
        }

        if (reference.IsSheetSpan)
            return ResolveSpan(reference, sheet);

        return ResolveOnSheet(reference, sheet);
    }

    private ReferenceValue? ResolveOnSheet(in FormulaRef reference, Worksheet sheet)
    {
        RangeRef? used = sheet.UsedRange;
        int lastUsedRow = used?.End.Row ?? 1;
        int lastUsedColumn = used?.End.Column ?? 1;

        (int firstRow, int firstColumn) = Corner(reference.Start, lastUsedRow, lastUsedColumn, start: true);
        (int lastRow, int lastColumn) = reference.IsRange
            ? Corner(reference.End, lastUsedRow, lastUsedColumn, start: false)
            : (firstRow, firstColumn);

        if (!SheetLimits.IsValidRow(firstRow) || !SheetLimits.IsValidRow(lastRow) ||
            !SheetLimits.IsValidColumn(firstColumn) || !SheetLimits.IsValidColumn(lastColumn))
        {
            return null;
        }

        return new ReferenceValue(sheet, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>
    /// The areas of a 3-D reference: every sheet from the first endpoint to the second in tab
    /// order, endpoints included, whichever way round they were written. Each sheet clamps
    /// whole columns to its own used range.
    /// </summary>
    private ReferenceValue? ResolveSpan(in FormulaRef reference, Worksheet firstSheet)
    {
        Worksheet? lastSheet = reference.SpanSheetId >= 0 ? Workbook.Sheets.FindById(reference.SpanSheetId) : null;
        if (lastSheet is null)
            return null;

        int firstIndex = TabIndex(firstSheet);
        int lastIndex = TabIndex(lastSheet);
        if (firstIndex < 0 || lastIndex < 0)
            return null;
        if (firstIndex > lastIndex)
            (firstIndex, lastIndex) = (lastIndex, firstIndex);

        ReferenceValue? head = null;
        ReferenceValue? tail = null;
        for (int index = firstIndex; index <= lastIndex; index++)
        {
            ReferenceValue? area = ResolveOnSheet(reference, Workbook.Sheets[index]);
            if (area is null)
                return null;

            if (tail is null)
                head = tail = area;
            else
                tail = ReferenceValue.Append(tail, area);
        }

        return head;
    }

    private int TabIndex(Worksheet sheet)
    {
        for (int i = 0; i < Workbook.Sheets.Count; i++)
        {
            if (ReferenceEquals(Workbook.Sheets[i], sheet))
                return i;
        }

        return -1;
    }

    /// <summary>
    /// Reads a reference into another workbook from the cached values its external link part
    /// carried. The cache is all there is: the engine never opens the other file, exactly like
    /// Excel with the source closed.
    /// </summary>
    public FormulaValue ResolveExternal(in FormulaRef reference)
    {
        IReadOnlyList<ExternalBook> books = Workbook.ExternalBooks;
        if (reference.ExternalBook > books.Count || reference.SheetNameIndex < 0)
            return FormulaValue.Error(ExcelError.Reference);

        // The books list is empty when the file carried no cache (or none was loaded); the
        // reference is then unanswerable rather than broken.
        if (books.Count == 0)
            return FormulaValue.Error(ExcelError.Reference);

        ExternalBook book = books[reference.ExternalBook - 1];
        return ExternalReader.Read(book, reference, this);
    }

    /// <summary>
    /// Turns a structured reference into a concrete area. Unlike an A1 reference this happens
    /// against the table's current shape, so a table that gained rows moves the reference with
    /// it. An unknown table or column, or a <c>[@…]</c> outside the table, yields nothing.
    /// </summary>
    public ReferenceValue? Resolve(in TableRef reference)
    {
        if (LocateTable(reference.TableName) is not var (sheet, table))
            return null;

        if (!TryRows(reference.Area, table, out int firstRow, out int lastRow))
            return null;
        if (!TryColumns(reference, table, out int firstColumn, out int lastColumn))
            return null;

        return new ReferenceValue(sheet, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>Finds a table by name, or the one holding the current cell when no name was given.</summary>
    private (Worksheet Sheet, ExcelTable Table)? LocateTable(string? name)
    {
        if (name is not null)
            return Workbook.LocateTable(name);

        foreach (ExcelTable table in Sheet.Tables)
        {
            if (table.Range.Contains(Cell))
                return (Sheet, table);
        }

        return null;
    }

    private bool TryRows(TableArea area, ExcelTable table, out int firstRow, out int lastRow)
    {
        int top = table.Range.Start.Row;
        int bottom = table.Range.End.Row;
        int dataFirst = table.HasHeaderRow ? top + 1 : top;
        int dataLast = table.HasTotalsRow ? bottom - 1 : bottom;

        if (area.HasFlag(TableArea.ThisRow))
        {
            firstRow = lastRow = Cell.Row;
            return Cell.Row >= top && Cell.Row <= bottom;
        }

        // Naming no band at all means the data, which is what Table[Column] is shorthand for.
        if (area == TableArea.None)
            area = TableArea.Data;

        firstRow = area.HasFlag(TableArea.Headers) && table.HasHeaderRow ? top : dataFirst;
        lastRow = area.HasFlag(TableArea.Totals) && table.HasTotalsRow ? bottom : dataLast;
        if (area == TableArea.Headers)
            (firstRow, lastRow) = (top, top);
        else if (area == TableArea.Totals)
            (firstRow, lastRow) = (bottom, bottom);

        return firstRow <= lastRow;
    }

    private static bool TryColumns(in TableRef reference, ExcelTable table, out int firstColumn, out int lastColumn)
    {
        firstColumn = table.Range.Start.Column;
        lastColumn = table.Range.End.Column;
        if (reference.FirstColumn is not { } first)
            return true;

        int offset = IndexOfColumn(table, first);
        if (offset < 0)
            return false;

        firstColumn = table.Range.Start.Column + offset;
        lastColumn = firstColumn;
        if (reference.LastColumn is not { } last)
            return true;

        int lastOffset = IndexOfColumn(table, last);
        if (lastOffset < 0)
            return false;

        lastColumn = table.Range.Start.Column + lastOffset;
        if (lastColumn < firstColumn)
            (firstColumn, lastColumn) = (lastColumn, firstColumn);
        return true;
    }

    private static int IndexOfColumn(ExcelTable table, string name)
    {
        for (int i = 0; i < table.Columns.Count; i++)
        {
            if (string.Equals(table.Columns[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }

    private static (int Row, int Column) Corner(RefPart part, int lastUsedRow, int lastUsedColumn, bool start) => part.Kind switch
    {
        RefPartKind.Cell => (part.Row, part.Column),
        RefPartKind.Column => (start ? 1 : lastUsedRow, part.Column),
        RefPartKind.Row => (part.Row, start ? 1 : lastUsedColumn),
        _ => (0, 0),
    };
}
