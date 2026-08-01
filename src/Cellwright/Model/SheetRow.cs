namespace Cellwright;

/// <summary>
/// Lightweight facade over one row of a <see cref="Worksheet"/>, in the spirit of
/// <see cref="Cell"/>: a (sheet, number) pair that reads and writes the sheet's row table.
/// Mutators return the same row for chaining: <c>sheet.Row(5).Hide().SetHeight(20)</c>.
/// </summary>
public readonly struct SheetRow
{
    private readonly Worksheet _sheet;

    internal SheetRow(Worksheet sheet, int number)
    {
        if (!SheetLimits.IsValidRow(number))
            throw new ArgumentOutOfRangeException(nameof(number), number, "Row numbers run from 1 to 1048576.");
        _sheet = sheet;
        Number = number;
    }

    /// <summary>The row's 1-based number.</summary>
    public int Number { get; }

    /// <summary>The row's stored properties; all defaults when the row was never configured.</summary>
    public RowProperties Properties => _sheet.RowTable.GetValueOrDefault(Number);

    /// <summary>Height in points, or <see langword="null"/> for the sheet default.</summary>
    public double? Height
    {
        get => Properties.Height;
        set => Update(Properties with { Height = Validated(value) });
    }

    /// <summary>Whether the row is hidden. Changing it makes SUBTOTAL and AGGREGATE recalculate.</summary>
    public bool Hidden
    {
        get => Properties.Hidden;
        set
        {
            if (Properties.Hidden == value)
                return;
            Update(Properties with { Hidden = value });
            _sheet.RaiseVisibilityChanged();
        }
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

    /// <summary>
    /// Style the row's cells inherit where they carry none of their own. One entry formats the
    /// whole row, which is what Excel does when you click a row header and pick a colour.
    /// </summary>
    public StyleId Style
    {
        get => Properties.Style;
        set => Update(Properties with { Style = value });
    }

    /// <summary>Sets the height in points.</summary>
    public SheetRow SetHeight(double height)
    {
        Height = height;
        return this;
    }

    /// <summary>Applies a registered style handle to the whole row.</summary>
    public SheetRow SetStyle(StyleId style)
    {
        Style = style;
        return this;
    }

    /// <summary>Registers and applies a style record to the whole row.</summary>
    public SheetRow SetStyle(CellStyle style) => SetStyle(_sheet.Workbook.RegisterStyle(style));

    /// <summary>Hides the row.</summary>
    public SheetRow Hide()
    {
        Hidden = true;
        return this;
    }

    /// <summary>Shows the row.</summary>
    public SheetRow Show()
    {
        Hidden = false;
        return this;
    }

    /// <summary>Returns a cell facade for a column of this row.</summary>
    public Cell Cell(int column) => _sheet.Cell(Number, column);

    private static double? Validated(double? height)
    {
        if (height is { } value)
            ArgumentOutOfRangeException.ThrowIfNegative(value);
        return height;
    }

    private static int ValidatedLevel(int level)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(level);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(level, 7);
        return level;
    }

    private void Update(RowProperties properties)
    {
        if (properties.IsDefault)
            _sheet.RowTable.Remove(Number);
        else
            _sheet.RowTable[Number] = properties;
    }
}

/// <summary>
/// A contiguous band of rows, returned by <see cref="Worksheet.Rows(int, int)"/>. Exists to
/// make grouping and visibility a single call instead of a loop at every call site.
/// </summary>
public readonly struct SheetRowRange
{
    private readonly Worksheet _sheet;

    internal SheetRowRange(Worksheet sheet, int first, int last)
    {
        if (!SheetLimits.IsValidRow(first) || !SheetLimits.IsValidRow(last))
            throw new ArgumentOutOfRangeException(nameof(first), "Row numbers run from 1 to 1048576.");
        _sheet = sheet;
        First = Math.Min(first, last);
        Last = Math.Max(first, last);
    }

    /// <summary>First row of the band.</summary>
    public int First { get; }

    /// <summary>Last row of the band, inclusive.</summary>
    public int Last { get; }

    /// <summary>Raises the outline level of every row in the band by one.</summary>
    public SheetRowRange Group(bool collapse = false)
    {
        for (int row = First; row <= Last; row++)
        {
            SheetRow target = _sheet.Row(row);
            target.OutlineLevel = Math.Min(target.OutlineLevel + 1, 7);
            if (collapse)
                target.Hidden = true;
        }

        if (collapse && SheetLimits.IsValidRow(Last + 1))
            _sheet.Row(Last + 1).Collapsed = true;
        return this;
    }

    /// <summary>Lowers the outline level of every row in the band by one.</summary>
    public SheetRowRange Ungroup()
    {
        for (int row = First; row <= Last; row++)
        {
            SheetRow target = _sheet.Row(row);
            target.OutlineLevel = Math.Max(target.OutlineLevel - 1, 0);
        }

        return this;
    }

    /// <summary>Hides every row in the band.</summary>
    public SheetRowRange Hide() => SetHidden(true);

    /// <summary>Shows every row in the band.</summary>
    public SheetRowRange Show() => SetHidden(false);

    /// <summary>Sets the height of every row in the band.</summary>
    public SheetRowRange SetHeight(double height)
    {
        for (int row = First; row <= Last; row++)
            _sheet.Row(row).Height = height;
        return this;
    }

    private SheetRowRange SetHidden(bool hidden)
    {
        for (int row = First; row <= Last; row++)
            _sheet.Row(row).Hidden = hidden;
        return this;
    }
}
