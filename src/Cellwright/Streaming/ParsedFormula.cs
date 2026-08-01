using Cellwright.IO;

namespace Cellwright.Streaming;

/// <summary>
/// The <c>f</c> element of a cell as it appears in the file: text plus the attributes that
/// tie shared and array formulas together.
/// </summary>
/// <param name="Text">Formula text, or <see langword="null"/> for a shared-formula member.</param>
/// <param name="Type">The <c>t</c> attribute: <c>shared</c>, <c>array</c>, <c>dataTable</c> or <see langword="null"/>.</param>
/// <param name="SharedIndex">The <c>si</c> attribute tying members to their master.</param>
/// <param name="Range">The <c>ref</c> attribute covering the group.</param>
/// <param name="Markup">Attributes the model does not represent, kept for the way back.</param>
internal readonly record struct ParsedFormula(
    string? Text, string? Type, int? SharedIndex, RangeRef? Range, FormulaMarkup? Markup = null)
{
    public bool IsShared => Type == "shared";

    public bool IsArray => Type == "array";

    /// <summary>Whether this is the master of a What-If data table (§18.3.1.40).</summary>
    public bool IsDataTable => Type == "dataTable";
}
