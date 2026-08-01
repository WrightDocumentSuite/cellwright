namespace Cellwright.Formulas;

/// <summary>Which line of a materialized rectangle a lookup searches.</summary>
internal enum LookupAxis
{
    /// <summary>The first column, as <c>VLOOKUP</c> searches it.</summary>
    FirstColumn = 0,

    /// <summary>The first row, as <c>HLOOKUP</c> searches it.</summary>
    FirstRow = 1,

    /// <summary>Every cell in reading order, as <c>MATCH</c> and <c>XLOOKUP</c> search it.</summary>
    AllValues = 2,
}

/// <summary>
/// Position of the first occurrence of each value along one axis of a range, so that an exact
/// lookup answers in constant time instead of scanning.
/// </summary>
/// <remarks>
/// The index has to agree with the linear scan exactly, and the comparison the scan uses is
/// only a well-behaved equivalence for numbers, text and logicals. Blanks and errors compare
/// equal to zero under those rules, which cannot be hashed consistently, so a line containing
/// either is left unindexed and the caller keeps scanning.
/// </remarks>
internal sealed class LookupIndex
{
    /// <summary>Shorter lines are scanned instead: building a dictionary would cost more.</summary>
    public const int MinLength = 64;

    private readonly Dictionary<double, int> _numbers = [];
    private readonly Dictionary<string, int> _texts = new(StringComparer.OrdinalIgnoreCase);
    private int _trueIndex = -1;
    private int _falseIndex = -1;

    private LookupIndex()
    {
    }

    /// <summary>
    /// Builds an index over one axis, or returns <see langword="null"/> when the line is too
    /// short or holds values that must fall back to scanning.
    /// </summary>
    public static LookupIndex? TryBuild(ArrayValue array, LookupAxis axis, ExcelDateSystem dateSystem)
    {
        int length = Length(array, axis);
        if (length < MinLength)
            return null;

        var index = new LookupIndex();
        for (int position = 0; position < length; position++)
        {
            if (!index.TryAdd(Value(array, axis, position), position, dateSystem))
                return null;
        }

        return index;
    }

    /// <summary>Whether a key can be answered from the index at all.</summary>
    public static bool CanSearch(in CellValue key) =>
        key.Kind is CellValueKind.Number or CellValueKind.Text or CellValueKind.Boolean
            or CellValueKind.DateTime or CellValueKind.TimeSpan;

    /// <summary>Finds the first position holding the key.</summary>
    public bool TryFind(in CellValue key, ExcelDateSystem dateSystem, out int position)
    {
        switch (key.Kind)
        {
            case CellValueKind.Text:
                return _texts.TryGetValue(key.GetText(), out position);

            case CellValueKind.Boolean:
                position = key.GetBoolean() ? _trueIndex : _falseIndex;
                return position >= 0;

            case CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan:
                return _numbers.TryGetValue(key.ToSerialNumber(dateSystem), out position);

            default:
                position = -1;
                return false;
        }
    }

    private bool TryAdd(in CellValue value, int position, ExcelDateSystem dateSystem)
    {
        switch (value.Kind)
        {
            case CellValueKind.Text:
                _texts.TryAdd(value.GetText(), position);
                return true;

            case CellValueKind.Boolean:
                if (value.GetBoolean())
                    _trueIndex = _trueIndex < 0 ? position : _trueIndex;
                else
                    _falseIndex = _falseIndex < 0 ? position : _falseIndex;
                return true;

            case CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan:
                _numbers.TryAdd(value.ToSerialNumber(dateSystem), position);
                return true;

            default:
                // Blanks and errors compare equal to zero in a scan; refuse to index the line.
                return false;
        }
    }

    private static int Length(ArrayValue array, LookupAxis axis) => axis switch
    {
        LookupAxis.FirstColumn => array.Rows,
        LookupAxis.FirstRow => array.Columns,
        _ => array.Count,
    };

    private static CellValue Value(ArrayValue array, LookupAxis axis, int position) => axis switch
    {
        LookupAxis.FirstColumn => array[position, 0],
        LookupAxis.FirstRow => array[0, position],
        _ => array.Values[position],
    };
}
