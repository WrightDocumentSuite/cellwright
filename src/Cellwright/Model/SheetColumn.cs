namespace Cellwright;

/// <summary>
/// Lightweight facade over one column of a <see cref="Worksheet"/>, the counterpart of
/// <see cref="SheetRow"/>. Mutators return the same column for chaining:
/// <c>sheet.Column("B").SetWidth(18).Hide()</c>.
/// </summary>
public readonly struct SheetColumn
{
    private readonly Worksheet _sheet;

    internal SheetColumn(Worksheet sheet, int number)
    {
        if (!SheetLimits.IsValidColumn(number))
            throw new ArgumentOutOfRangeException(nameof(number), number, "Column numbers run from 1 to 16384.");
        _sheet = sheet;
        Number = number;
    }

    /// <summary>The sheet the column belongs to.</summary>
    internal Worksheet Sheet => _sheet;

    /// <summary>The column's 1-based number.</summary>
    public int Number { get; }

    /// <summary>The column's letter name, e.g. <c>AB</c>.</summary>
    public string Name => ColumnName.Format(Number);

    /// <summary>The column's stored properties; all defaults when the column was never configured.</summary>
    public ColumnProperties Properties => _sheet.ColumnTable.GetValueOrDefault(Number);

    /// <summary>Width in characters, or <see langword="null"/> for the sheet default.</summary>
    public double? Width
    {
        get => Properties.Width;
        set => Update(Properties with { Width = Validated(value) });
    }

    /// <summary>Whether the column is hidden.</summary>
    public bool Hidden
    {
        get => Properties.Hidden;
        set => Update(Properties with { Hidden = value });
    }

    /// <summary>Outline (grouping) level, 0 through 7.</summary>
    public int OutlineLevel
    {
        get => Properties.OutlineLevel;
        set => Update(Properties with { OutlineLevel = ValidatedLevel(value) });
    }

    /// <summary>Whether the outline group ending here is shown collapsed.</summary>
    public bool Collapsed
    {
        get => Properties.Collapsed;
        set => Update(Properties with { Collapsed = value });
    }

    /// <summary>Style the column's cells inherit where they carry none of their own.</summary>
    public StyleId Style
    {
        get => Properties.Style;
        set => Update(Properties with { Style = value });
    }

    /// <summary>Sets the width in characters.</summary>
    public SheetColumn SetWidth(double width)
    {
        Width = width;
        return this;
    }

    /// <summary>Applies a registered style handle to the whole column.</summary>
    public SheetColumn SetStyle(StyleId style)
    {
        Style = style;
        return this;
    }

    /// <summary>Registers and applies a style record to the whole column.</summary>
    public SheetColumn SetStyle(CellStyle style) => SetStyle(_sheet.Workbook.RegisterStyle(style));

    /// <summary>Hides the column.</summary>
    public SheetColumn Hide()
    {
        Hidden = true;
        return this;
    }

    /// <summary>Shows the column.</summary>
    public SheetColumn Show()
    {
        Hidden = false;
        return this;
    }

    /// <summary>Returns a cell facade for a row of this column.</summary>
    public Cell Cell(int row) => _sheet.Cell(row, Number);

    private static double? Validated(double? width)
    {
        if (width is { } value)
            ArgumentOutOfRangeException.ThrowIfNegative(value);
        return width;
    }

    private static int ValidatedLevel(int level)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(level);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(level, 7);
        return level;
    }

    private void Update(ColumnProperties properties)
    {
        if (properties.IsDefault)
            _sheet.ColumnTable.Remove(Number);
        else
            _sheet.ColumnTable[Number] = properties;
    }
}

/// <summary>
/// A contiguous band of columns, returned by <see cref="Worksheet.Columns(int, int)"/>.
/// </summary>
public readonly struct SheetColumnRange
{
    private readonly Worksheet _sheet;

    internal SheetColumnRange(Worksheet sheet, int first, int last)
    {
        if (!SheetLimits.IsValidColumn(first) || !SheetLimits.IsValidColumn(last))
            throw new ArgumentOutOfRangeException(nameof(first), "Column numbers run from 1 to 16384.");
        _sheet = sheet;
        First = Math.Min(first, last);
        Last = Math.Max(first, last);
    }

    /// <summary>First column of the band.</summary>
    public int First { get; }

    /// <summary>Last column of the band, inclusive.</summary>
    public int Last { get; }

    /// <summary>Raises the outline level of every column in the band by one.</summary>
    public SheetColumnRange Group(bool collapse = false)
    {
        for (int column = First; column <= Last; column++)
        {
            SheetColumn target = _sheet.Column(column);
            target.OutlineLevel = Math.Min(target.OutlineLevel + 1, 7);
            if (collapse)
                target.Hidden = true;
        }

        if (collapse && SheetLimits.IsValidColumn(Last + 1))
            _sheet.Column(Last + 1).Collapsed = true;
        return this;
    }

    /// <summary>Lowers the outline level of every column in the band by one.</summary>
    public SheetColumnRange Ungroup()
    {
        for (int column = First; column <= Last; column++)
        {
            SheetColumn target = _sheet.Column(column);
            target.OutlineLevel = Math.Max(target.OutlineLevel - 1, 0);
        }

        return this;
    }

    /// <summary>Hides every column in the band.</summary>
    public SheetColumnRange Hide() => SetHidden(true);

    /// <summary>Shows every column in the band.</summary>
    public SheetColumnRange Show() => SetHidden(false);

    /// <summary>Sets the width of every column in the band.</summary>
    public SheetColumnRange SetWidth(double width)
    {
        for (int column = First; column <= Last; column++)
            _sheet.Column(column).Width = width;
        return this;
    }

    private SheetColumnRange SetHidden(bool hidden)
    {
        for (int column = First; column <= Last; column++)
            _sheet.Column(column).Hidden = hidden;
        return this;
    }
}
