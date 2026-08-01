using Cellwright.IO;
using Cellwright.Streaming;

namespace Cellwright;

/// <summary>
/// Lightweight facade over one cell of a <see cref="Worksheet"/>: a (sheet, reference) pair
/// that reads and writes the sheet's column slices directly. Mutators return the same cell
/// for chaining: <c>sheet.Cell("A1").SetValue(1).SetStyle(header)</c>.
/// </summary>
public readonly struct Cell
{
    private readonly Worksheet _sheet;

    internal Cell(Worksheet sheet, CellRef reference)
    {
        _sheet = sheet;
        Ref = reference;
    }

    internal Worksheet Sheet => _sheet;

    /// <summary>The cell's reference.</summary>
    public CellRef Ref { get; }

    /// <summary>The cell's value; blank when unset.</summary>
    public CellValue Value => _sheet.Values.Get(Ref.Row, Ref.Column);

    /// <summary>Formula text without the leading <c>=</c>, or <see langword="null"/>.</summary>
    public string? Formula => _sheet.Formulas.Get(Ref.Row, Ref.Column);

    /// <summary>The style handle applied to the cell.</summary>
    public StyleId Style => new(_sheet.StyleIndexes.Get(Ref.Row, Ref.Column));

    /// <summary>The full style record applied to the cell.</summary>
    public CellStyle StyleValue => _sheet.Workbook.GetStyle(Style);

    /// <summary>
    /// The rectangle this cell's dynamic array formula spilled into, or <see langword="null"/>
    /// when it is not a spill anchor. Filled by <c>Calculate()</c> from the Cellwright.Formulas
    /// package; a workbook that was never calculated has no spills.
    /// </summary>
    public RangeRef? SpillRange => _sheet.SpillTable.TryGetValue(Ref, out RangeRef range) ? range : null;

    /// <summary>
    /// Whether this cell's value was spilled into it by a dynamic array formula anchored
    /// somewhere else. Such a cell has a value but no formula of its own.
    /// </summary>
    public bool IsSpilled => _sheet.SpillAnchorCovering(Ref) is { } anchor && anchor != Ref;

    /// <summary>
    /// The cell's text with the character formatting inside it, or <see langword="null"/> when it
    /// carries none.
    /// </summary>
    /// <remarks>
    /// Formatting belongs to the text it was applied to. A cell whose value has been replaced
    /// since the runs were read answers <see langword="null"/> rather than handing back
    /// formatting that was measured against words no longer there.
    /// </remarks>
    public RichText? RichText
    {
        get
        {
            if (!_sheet.RichStringTable.TryGetValue(Ref, out RichStringMarkup stored))
                return null;

            CellValue value = Value;
            if (!value.IsText || !string.Equals(value.GetText(), stored.Text, StringComparison.Ordinal))
                return null;

            return RichTextMarkup.Parse(stored.Markup);
        }
    }

    /// <summary>
    /// Sets the cell's text and the character formatting inside it. An existing formula is
    /// removed, as with any other value.
    /// </summary>
    /// <param name="text">The runs to write.</param>
    /// <remarks>
    /// A text whose runs all follow the cell's own font is stored as an ordinary string: there
    /// would be nothing for the runs to say.
    /// </remarks>
    public Cell SetRichText(RichText text)
    {
        ArgumentNullException.ThrowIfNull(text);

        bool hadFormula = _sheet.Formulas.Get(Ref.Row, Ref.Column) is not null;
        string plain = text.Text;

        _sheet.Values.Set(Ref.Row, Ref.Column, CellValue.Text(plain));
        _sheet.Formulas.Set(Ref.Row, Ref.Column, null);

        if (text.IsFormatted)
            _sheet.RichStringTable[Ref] = new RichStringMarkup(plain, RichTextMarkup.ToMarkup(text));
        else
            _sheet.RichStringTable.Remove(Ref);

        _sheet.RaiseCellChanged(Ref, structural: hadFormula);
        return this;
    }

    /// <summary>Sets the value. An existing formula is removed, mirroring Excel's behavior.</summary>
    public Cell SetValue(CellValue value)
    {
        bool hadFormula = _sheet.Formulas.Get(Ref.Row, Ref.Column) is not null;
        _sheet.Values.Set(Ref.Row, Ref.Column, value);
        _sheet.Formulas.Set(Ref.Row, Ref.Column, null);

        // The runs described the text that was here; a new value is not that text.
        _sheet.RichStringTable.Remove(Ref);

        _sheet.RaiseCellChanged(Ref, structural: hadFormula);
        return this;
    }

    /// <summary>
    /// Sets the formula text (without the leading <c>=</c>) and clears the cached value.
    /// Reference the Cellwright.Formulas package and call <c>Calculate()</c> to compute it;
    /// otherwise Excel recalculates when the file is opened.
    /// </summary>
    public Cell SetFormula(string? formula)
    {
        if (formula is not null && formula.StartsWith('='))
            formula = formula[1..];
        _sheet.Formulas.Set(Ref.Row, Ref.Column, string.IsNullOrEmpty(formula) ? null : formula);
        _sheet.Values.Set(Ref.Row, Ref.Column, CellValue.Blank);
        _sheet.RaiseCellChanged(Ref, structural: true);
        return this;
    }

    /// <summary>
    /// Stores a computed result without touching the formula and without marking the cell
    /// dirty. Used by the calculation engine to publish results into the value slice.
    /// </summary>
    internal void SetComputedValue(CellValue value) => _sheet.Values.Set(Ref.Row, Ref.Column, value);

    /// <summary>Applies a registered style handle.</summary>
    public Cell SetStyle(StyleId style)
    {
        _sheet.StyleIndexes.Set(Ref.Row, Ref.Column, style.Index);
        return this;
    }

    /// <summary>Registers and applies a style record.</summary>
    public Cell SetStyle(CellStyle style) => SetStyle(_sheet.Workbook.RegisterStyle(style));

    /// <summary>Clears value, formula and style.</summary>
    public Cell Clear()
    {
        _sheet.Values.Set(Ref.Row, Ref.Column, CellValue.Blank);
        _sheet.Formulas.Set(Ref.Row, Ref.Column, null);
        _sheet.StyleIndexes.Set(Ref.Row, Ref.Column, 0);
        _sheet.RichStringTable.Remove(Ref);
        _sheet.RaiseCellChanged(Ref, structural: true);
        return this;
    }

    /// <inheritdoc/>
    public override string ToString() => $"{Ref}: {Value}";
}
