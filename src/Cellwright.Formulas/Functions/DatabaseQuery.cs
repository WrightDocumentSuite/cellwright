namespace Cellwright.Formulas.Functions;

/// <summary>
/// The table shape the database functions share: records under a row of column names, and a
/// second table of conditions written the same way.
/// </summary>
/// <remarks>
/// Conditions side by side on one row must all hold; conditions on separate rows are
/// alternatives. That layout is the entire language of these functions — there is no other way
/// to say "and" or "or" to them, which is why one criteria table can be reused for several
/// queries and why a column of empty conditions has to mean "no condition" rather than "empty".
/// </remarks>
internal readonly struct DatabaseQuery
{
    private readonly ArrayValue _database;
    private readonly int _criteriaRows;
    private readonly int _criteriaColumns;
    private readonly Condition[] _conditions;

    private DatabaseQuery(ArrayValue database, ArrayValue criteria, Condition[] conditions)
    {
        _database = database;
        _criteriaRows = criteria.Rows;
        _criteriaColumns = criteria.Columns;
        _conditions = conditions;
    }

    /// <summary>One condition cell, already interpreted.</summary>
    private readonly record struct Condition(bool Stated, int Column, Criterion Test);

    /// <summary>How many records the database holds, its naming row aside.</summary>
    public int Records => _database.Rows - 1;

    /// <summary>
    /// Pairs each condition with the database column it names and interprets it once, rather
    /// than re-reading the criteria table for every record.
    /// </summary>
    public static bool TryBuild(
        ArrayValue database, ArrayValue criteria, EvalContext context, out DatabaseQuery query)
    {
        query = default;
        if (database.Rows < 2 || criteria.Rows < 1 || criteria.Columns < 1)
            return false;

        Span<int> targets = criteria.Columns <= 32 ? stackalloc int[criteria.Columns] : new int[criteria.Columns];
        for (int column = 0; column < criteria.Columns; column++)
        {
            // A column stating nothing needs no name, so its heading is never looked up. That
            // is what lets a criteria range be wider than the conditions written into it.
            if (IsSilent(criteria, column))
            {
                targets[column] = -1;
                continue;
            }

            if (!TryFindColumn(database, criteria[0, column], context, out int target))
                return false;

            targets[column] = target;
        }

        query = new DatabaseQuery(database, criteria, Interpret(criteria, targets, context));
        return true;
    }

    /// <summary>Whether the conditions admit the record at <paramref name="record"/>, counted from zero.</summary>
    public bool Admits(int record, EvalContext context)
    {
        // A criteria table that is nothing but its naming row states no condition at all.
        if (_criteriaRows < 2)
            return true;

        for (int row = 1; row < _criteriaRows; row++)
        {
            if (RowAdmits(row, record, context))
                return true;
        }

        return false;
    }

    /// <summary>The value one record holds in one column.</summary>
    public CellValue Value(int record, int column) => _database[record + 1, column];

    /// <summary>
    /// Resolves the field argument: the name of a column, or its position counted from one.
    /// </summary>
    public static bool TryResolveField(
        ArrayValue database, in CellValue field, EvalContext context, out int column)
    {
        if (field.IsNumber)
        {
            column = (int)field.GetNumber() - 1;
            return column >= 0 && column < database.Columns;
        }

        return TryFindColumn(database, field, context, out column);
    }

    private static Condition[] Interpret(ArrayValue criteria, ReadOnlySpan<int> targets, EvalContext context)
    {
        var conditions = new Condition[(criteria.Rows - 1) * criteria.Columns];
        for (int row = 1; row < criteria.Rows; row++)
        {
            for (int column = 0; column < criteria.Columns; column++)
            {
                CellValue cell = criteria[row, column];
                if (targets[column] < 0 || cell.IsBlank)
                    continue;

                conditions[((row - 1) * criteria.Columns) + column] =
                    new Condition(true, targets[column], Criterion.Parse(cell, context));
            }
        }

        return conditions;
    }

    private bool RowAdmits(int row, int record, EvalContext context)
    {
        int start = (row - 1) * _criteriaColumns;
        for (int column = 0; column < _criteriaColumns; column++)
        {
            Condition condition = _conditions[start + column];
            if (condition.Stated && !condition.Test.Matches(_database[record + 1, condition.Column], context))
                return false;
        }

        return true;
    }

    private static bool IsSilent(ArrayValue criteria, int column)
    {
        for (int row = 1; row < criteria.Rows; row++)
        {
            if (!criteria[row, column].IsBlank)
                return false;
        }

        return true;
    }

    private static bool TryFindColumn(
        ArrayValue database, in CellValue name, EvalContext context, out int column)
    {
        for (int candidate = 0; candidate < database.Columns; candidate++)
        {
            if (Coercion.Compare(database[0, candidate], name, context.DateSystem) == 0)
            {
                column = candidate;
                return true;
            }
        }

        column = -1;
        return false;
    }
}
