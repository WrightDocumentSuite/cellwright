namespace Cellwright.IO;

/// <summary>
/// An entry of the workbook's sheet list that is not a worksheet — a chart sheet, a dialog
/// sheet or a macro sheet. There is no grid to model, so the entry and its part travel
/// through untouched at the position they held.
/// </summary>
/// <param name="Position">Zero-based index within the original sheet list.</param>
/// <param name="Name">Tab name.</param>
/// <param name="State">The <c>state</c> attribute, or <see langword="null"/> when visible.</param>
/// <param name="RelationshipId">Relationship id pointing at the sheet's part.</param>
/// <param name="SheetId">The <c>sheetId</c> the workbook gave it, which other parts refer to.</param>
internal readonly record struct ForeignSheet(
    int Position, string Name, string? State, string RelationshipId, int SheetId);
