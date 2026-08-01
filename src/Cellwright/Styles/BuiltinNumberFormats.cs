using System.Collections.Frozen;

namespace Cellwright.Styles;

/// <summary>
/// The builtin number formats of SpreadsheetML: ids 0–49 have implicit format codes
/// that are never written to the styles part.
/// </summary>
internal static class BuiltinNumberFormats
{
    /// <summary>First id available for custom formats.</summary>
    public const int FirstCustomId = 164;

    private static readonly FrozenDictionary<int, string> CodeById = new Dictionary<int, string>
    {
        [0] = "",
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "m/d/yyyy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    }.ToFrozenDictionary();

    private static readonly FrozenDictionary<string, int> IdByCode =
        CodeById.ToFrozenDictionary(static pair => pair.Value, static pair => pair.Key, StringComparer.Ordinal);

    /// <summary>Returns the implicit code of a builtin id, or <see langword="null"/>.</summary>
    public static string? GetCode(int id) => CodeById.TryGetValue(id, out string? code) ? code : null;

    /// <summary>Returns the builtin id matching a code exactly, or <see langword="null"/>.</summary>
    public static int? GetId(string code) => IdByCode.TryGetValue(code, out int id) ? id : null;
}
