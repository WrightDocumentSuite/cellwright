namespace Cellwright;

/// <summary>
/// How a group of cells shares one formula in the file format.
/// </summary>
public enum FormulaGroupKind
{
    /// <summary>
    /// A shared formula: the master cell holds the text, members repeat it with
    /// relative references shifted by their offset from the master.
    /// </summary>
    Shared,

    /// <summary>An array formula entered over a range (legacy CSE formula).</summary>
    Array,
}

/// <summary>
/// A shared or array formula as stored in the file. The core preserves groups as data;
/// the Cellwright.Formulas package expands shared groups into per-cell formulas.
/// </summary>
public sealed class FormulaGroup
{
    internal FormulaGroup(FormulaGroupKind kind, CellRef master, RangeRef range, string text)
    {
        Kind = kind;
        Master = master;
        Range = range;
        Text = text;
    }

    /// <summary>Shared or array.</summary>
    public FormulaGroupKind Kind { get; }

    /// <summary>Cell holding the formula text.</summary>
    public CellRef Master { get; }

    /// <summary>
    /// Range covered by the group. Grows while loading if members turn up outside the
    /// range declared by the master.
    /// </summary>
    public RangeRef Range { get; internal set; }

    /// <summary>Formula text as written at the master cell, without the leading <c>=</c>.</summary>
    public string Text { get; }

    /// <inheritdoc/>
    public override string ToString() => $"{Kind} {Range} = {Text}";
}
