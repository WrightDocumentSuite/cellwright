namespace Cellwright.Formulas.Parsing;

/// <summary>
/// A reference operand of a compiled formula. Coordinates are absolute as written at the
/// formula's own cell; the anchor flags say which of them follow the cell when the formula
/// is copied, which is what shared-formula expansion needs.
/// </summary>
/// <param name="SheetId">
/// Identity of the target sheet, <see cref="CurrentSheet"/> for the formula's own sheet, or
/// <see cref="UnresolvedSheet"/> when the name did not match any sheet.
/// </param>
/// <param name="SheetNameIndex">Index into the string pool holding the original sheet name.</param>
/// <param name="Start">Top-left part of the reference.</param>
/// <param name="End">Bottom-right part, or <see cref="RefPartKind.None"/> for a single cell.</param>
/// <param name="SpanSheetId">Identity of the last sheet of a 3-D span, or <see cref="NoSheet"/>.</param>
/// <param name="SpanSheetNameIndex">String-pool index of the span's last sheet name, or -1.</param>
/// <param name="ExternalBook">One-based external book index of a <c>[N]Sheet!</c> prefix, or 0.</param>
/// <param name="ExternalSheetName">
/// Sheet name of an external reference, carried by value: the cached book resolves names at
/// evaluation, when the program's string pool is out of reach.
/// </param>
/// <param name="ExternalSpanEnd">Last sheet of an external 3-D span, or <see langword="null"/>.</param>
/// <param name="IsSpilled">
/// Whether the reference carries the spill suffix (<c>A1#</c>): it names whatever rectangle the
/// dynamic array anchored at the cell spilled into, not the cell itself.
/// </param>
internal readonly record struct FormulaRef(
    int SheetId, int SheetNameIndex, RefPart Start, RefPart End,
    int SpanSheetId = FormulaRef.NoSheet, int SpanSheetNameIndex = -1, int ExternalBook = 0,
    string? ExternalSheetName = null, string? ExternalSpanEnd = null, bool IsSpilled = false)
{
    /// <summary>Sheet id meaning "the sheet the formula lives on".</summary>
    public const int CurrentSheet = -1;

    /// <summary>Sheet id meaning "a sheet name that does not exist"; evaluates to <c>#REF!</c>.</summary>
    public const int UnresolvedSheet = -2;

    /// <summary>Span sheet id meaning "no 3-D span at all".</summary>
    public const int NoSheet = -3;

    /// <summary>Whether the reference runs across a span of sheets.</summary>
    public bool IsSheetSpan => SpanSheetId != NoSheet;

    /// <summary>Whether the reference reads another workbook's cached cells.</summary>
    public bool IsExternal => ExternalBook > 0;

    /// <summary>Whether the reference spans more than one cell.</summary>
    public bool IsRange => End.Kind != RefPartKind.None;

    /// <summary>Whether the reference covers entire columns, e.g. <c>A:C</c>.</summary>
    public bool IsWholeColumn => Start.Kind == RefPartKind.Column;

    /// <summary>Whether the reference covers entire rows, e.g. <c>2:5</c>.</summary>
    public bool IsWholeRow => Start.Kind == RefPartKind.Row;

    /// <summary>
    /// Returns the reference as it would read after copying the formula by the given offset:
    /// relative coordinates move, anchored ones stay.
    /// </summary>
    public FormulaRef Shift(int rowDelta, int columnDelta) =>
        this with { Start = ShiftPart(Start, rowDelta, columnDelta), End = ShiftPart(End, rowDelta, columnDelta) };

    private static RefPart ShiftPart(RefPart part, int rowDelta, int columnDelta) => part.Kind switch
    {
        RefPartKind.None => part,
        _ => part with
        {
            Row = part.RowAbsolute || part.Kind == RefPartKind.Column ? part.Row : part.Row + rowDelta,
            Column = part.ColumnAbsolute || part.Kind == RefPartKind.Row ? part.Column : part.Column + columnDelta,
        },
    };
}
