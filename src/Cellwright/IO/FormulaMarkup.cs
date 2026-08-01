namespace Cellwright.IO;

/// <summary>
/// Attributes of a cell's <c>f</c> element that the model does not represent, kept so that
/// saving puts them back.
/// </summary>
/// <remarks>
/// Two cases arrive here. Ordinary formulas may carry flags such as <c>ca</c> — "always
/// recalculate", which is how a volatile add-in function stays fresh — and those are spliced
/// into the element the writer generates. A What-If data table (§18.3.1.40) is the other case:
/// its master has no formula text at all, only attributes describing the two input cells, so
/// the whole element is reproduced and the model stays out of it.
/// </remarks>
/// <param name="Attributes">Attributes to write, name and value as read.</param>
/// <param name="WritesOwnElement">Whether the attributes describe the whole element rather than additions to it.</param>
internal readonly record struct FormulaMarkup(
    List<KeyValuePair<string, string>> Attributes, bool WritesOwnElement);
