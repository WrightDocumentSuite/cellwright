using Cellwright.Model;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Turns the workbook's Lbl records into defined names. Some of them are not names a person
/// wrote: Excel's own <c>Print_Area</c>, <c>Print_Titles</c> and <c>_FilterDatabase</c> are how
/// the format states settings that the model keeps elsewhere, so they are applied as well.
/// </summary>
internal static class DefinedNameLoader
{
    /// <summary>Applies every name to the workbook.</summary>
    public static void Apply(Workbook workbook, XlsGlobals globals, PtgDecoder decoder, Action<WorkbookWarning> warn)
    {
        // The name's sheet index counts the sheets the file declares, including the chart and
        // macro ones the loader skipped, so the two orders have to be reconciled.
        int[] sheetMap = BuildSheetMap(workbook, globals);

        foreach (XlsDefinedName name in globals.Names)
        {
            string? refersTo = decoder.Decode(name.Rgce, new CellRef(1, 1), sheetIndex: 0);
            if (refersTo is null)
            {
                warn(new WorkbookWarning(
                    "UndecodedName",
                    $"Defined name '{name.Name}' refers to something this library cannot render and was dropped."));
                continue;
            }

            Worksheet? scope = ResolveScope(workbook, sheetMap, name.SheetIndex);
            if (name.SheetIndex > 0 && scope is null)
                continue;

            if (ApplyBuiltIn(name, refersTo, scope))
                continue;

            try
            {
                DefinedName defined = workbook.Names.Add(name.Name, refersTo, scope);
                defined.IsHidden = name.IsHidden;
            }
            catch (ArgumentException)
            {
                warn(new WorkbookWarning(
                    "BadDefinedName",
                    $"Defined name '{name.Name}' is not one this library can represent and was dropped."));
            }
        }
    }

    /// <summary>
    /// Applies one of Excel's own names to the model instead of adding it as a name. Returns
    /// whether it was handled.
    /// </summary>
    private static bool ApplyBuiltIn(in XlsDefinedName name, string refersTo, Worksheet? scope)
    {
        if (!name.IsBuiltIn || scope is null)
            return false;

        switch (name.Name)
        {
            case "_xlnm.Print_Area":
                if (RangeRef.TryParse(StripSheet(refersTo), null, out RangeRef area))
                    scope.PrintArea = area;
                return true;

            case "_xlnm._FilterDatabase":
                if (RangeRef.TryParse(StripSheet(refersTo), null, out RangeRef range))
                {
                    scope.AutoFilterRange = range;
                    scope.AutoFilterIsActive = true;
                }

                return true;

            default:
                return false;
        }
    }

    /// <summary>Drops the sheet prefix of a reference, leaving the range on its own.</summary>
    private static string StripSheet(string reference)
    {
        int bang = reference.LastIndexOf('!');
        return bang < 0 ? reference : reference[(bang + 1)..];
    }

    private static Worksheet? ResolveScope(Workbook workbook, int[] sheetMap, int oneBasedSheetIndex)
    {
        if (oneBasedSheetIndex <= 0)
            return null;

        int declared = oneBasedSheetIndex - 1;
        if ((uint)declared >= (uint)sheetMap.Length)
            return null;

        int loaded = sheetMap[declared];
        return (uint)loaded < (uint)workbook.Sheets.Count ? workbook.Sheets[loaded] : null;
    }

    /// <summary>Maps each declared sheet position onto its position in the loaded workbook.</summary>
    private static int[] BuildSheetMap(Workbook workbook, XlsGlobals globals)
    {
        var map = new int[globals.Sheets.Count];
        int loaded = 0;
        for (int i = 0; i < globals.Sheets.Count; i++)
            map[i] = globals.Sheets[i].Kind == XlsSheetKind.Worksheet ? loaded++ : -1;

        _ = workbook;
        return map;
    }
}
