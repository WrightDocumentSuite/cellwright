namespace Cellwright.Formulas.Functions;

/// <summary>
/// <c>CELL</c> and <c>INFO</c> — the functions that describe the workbook rather than compute
/// on it. Both are volatile: what they report (formatting, sheet count) changes without any
/// cell being written, so the dependency graph cannot see it.
/// </summary>
internal static class CellInformationFunctions
{
    public static void Register(FunctionRegistryBuilder builder)
    {
        builder.Add("CELL", 1, 2, Cell, isVolatile: true, wantsRawReferences: true);
        builder.Add("INFO", 1, 1, Info, isVolatile: true);
    }

    private static FormulaValue Cell(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string infoType = Args.Text(context, args[0]).ToLowerInvariant();

        // Without a reference Excel reports on the last-changed cell; an engine that
        // recalculates in graph order has no such notion, so the reference is required.
        if (args.Length < 2 || args[1].Kind != FormulaValueKind.Reference)
            return FormulaValue.Error(ExcelError.Value);

        ReferenceValue reference = args[1].AsReference;
        Worksheet sheet = reference.Sheet;
        var cell = new CellRef(reference.FirstRow, reference.FirstColumn);
        CellStyle style = sheet.Cell(cell).StyleValue;

        switch (infoType)
        {
            case "address":
                return FormulaValue.Text($"${ColumnName.Format(cell.Column)}${cell.Row}");
            case "col":
                return FormulaValue.Number(cell.Column);
            case "row":
                return FormulaValue.Number(cell.Row);
            case "contents":
                return FormulaValue.Scalar(context.ReadCell(sheet, cell.Row, cell.Column));
            case "type":
            {
                CellValue value = context.ReadCell(sheet, cell.Row, cell.Column);
                return FormulaValue.Text(value.IsBlank ? "b" : value.IsText ? "l" : "v");
            }

            case "protect":
                return FormulaValue.Number(style.Protection.Locked ? 1 : 0);
            case "width":
                return FormulaValue.Number(Math.Truncate(sheet.Column(cell.Column).Width ?? 8.43));
            case "prefix":
            {
                CellValue value = context.ReadCell(sheet, cell.Row, cell.Column);
                if (!value.IsText)
                    return FormulaValue.Text("");
                return FormulaValue.Text(style.Alignment.Horizontal switch
                {
                    HorizontalAlignment.Right => "\"",
                    HorizontalAlignment.Center => "^",
                    HorizontalAlignment.Fill => "\\",
                    _ => "'",
                });
            }

            case "filename":
                // The engine works on a model, not a file on disk; Excel returns "" for a
                // book that has never been saved, and that is the honest answer here too.
                return FormulaValue.Text("");
            case "color":
            case "parentheses":
                return FormulaValue.Number(FormatSections(style.NumberFormat, infoType));
            case "format":
                return FormulaValue.Text(FormatCode(style.NumberFormat));
            default:
                return FormulaValue.Error(ExcelError.Value);
        }
    }

    /// <summary>Whether the negative section colors or parenthesizes, as CELL reports them.</summary>
    private static int FormatSections(string code, string infoType)
    {
        int semicolon = IndexOfSectionBreak(code);
        if (semicolon < 0)
            return 0;

        string negative = code[(semicolon + 1)..];
        return infoType == "color"
            ? negative.Contains('[') ? 1 : 0
            : negative.Contains('(') ? 1 : 0;
    }

    private static int IndexOfSectionBreak(string code)
    {
        bool quoted = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '"')
                quoted = !quoted;
            else if (c == ';' && !quoted)
                return i;
        }

        return -1;
    }

    /// <summary>
    /// The terse format classes of <c>CELL("format")</c>: G for general, F and comma classes
    /// with their decimal count, C for currency, P for percent, S for scientific, D1–D9 for
    /// the date and time shapes.
    /// </summary>
    internal static string FormatCode(string code)
    {
        if (code.Length == 0 || code == "General")
            return "G";

        // Only the first (positive) section decides the class.
        int semicolon = IndexOfSectionBreak(code);
        string section = semicolon < 0 ? code : code[..semicolon];
        string bare = StripLiterals(section);

        if (bare.Contains('%'))
            return "P" + Decimals(bare);
        if (bare.Contains('E') || bare.Contains('e'))
            return "S" + Decimals(bare);

        bool temporal = bare.IndexOfAny(['y', 'd', 'h', 's']) >= 0 ||
            (bare.Contains('m') && !bare.Contains('#') && !bare.Contains('0'));
        if (temporal)
            return DateCode(bare);

        if (section.Contains('$') || section.Contains('¤'))
            return "C" + Decimals(bare);
        if (bare.Contains("#,#") || bare.Contains("0,0"))
            return "," + Decimals(bare);
        if (bare.Contains('0') || bare.Contains('#'))
            return "F" + Decimals(bare);

        return "G";
    }

    private static string StripLiterals(string section)
    {
        var builder = new System.Text.StringBuilder(section.Length);
        bool quoted = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '"')
            {
                quoted = !quoted;
                continue;
            }

            if (quoted)
                continue;
            if (c == '\\' && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c == '[')
            {
                int close = section.IndexOf(']', i);
                i = close < 0 ? section.Length : close;
                continue;
            }

            builder.Append(c);
        }

        return builder.ToString();
    }

    private static int Decimals(string bare)
    {
        int point = bare.IndexOf('.');
        if (point < 0)
            return 0;

        int count = 0;
        for (int i = point + 1; i < bare.Length && bare[i] is '0' or '#' or '?'; i++)
            count++;
        return count;
    }

    private static string DateCode(string bare)
    {
        bool hasDate = bare.IndexOfAny(['y', 'd']) >= 0 ||
            (bare.Contains('m') && bare.IndexOfAny(['h', 's']) < 0);
        bool hasTime = bare.IndexOfAny(['h', 's']) >= 0;

        if (hasDate)
        {
            if (hasTime)
                return "D4";
            if (bare.Contains('y') && bare.Contains("mmm", StringComparison.Ordinal))
                return bare.Contains('d') ? "D1" : "D3";
            if (bare.Contains("mmm", StringComparison.Ordinal))
                return "D2";
            if (!bare.Contains('y'))
                return "D5";
            return "D4";
        }

        bool seconds = bare.Contains('s');
        bool twelveHour = bare.Contains("AM/PM", StringComparison.OrdinalIgnoreCase) ||
            bare.Contains("A/P", StringComparison.OrdinalIgnoreCase);
        return (seconds, twelveHour) switch
        {
            (true, true) => "D6",
            (false, true) => "D7",
            (true, false) => "D8",
            _ => "D9",
        };
    }

    private static FormulaValue Info(EvalContext context, ReadOnlySpan<FormulaValue> args)
    {
        string type = Args.Text(context, args[0]).ToLowerInvariant();
        return type switch
        {
            "numfile" => FormulaValue.Number(context.Workbook.Sheets.Count),
            "recalc" => FormulaValue.Text("Manual"),
            "system" => FormulaValue.Text("pcdos"),

            // Directory, OS version and release describe a machine and an application this
            // engine is neither of; Excel's own answer for something it cannot know is #N/A.
            "directory" or "osversion" or "release" or "origin" or "memavail" or "memused" or "totmem"
                => FormulaValue.Error(ExcelError.NotAvailable),
            _ => FormulaValue.Error(ExcelError.Value),
        };
    }
}
