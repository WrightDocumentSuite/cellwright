using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;

namespace Cellwright.Xls.Writing;

/// <summary>
/// Writes the input rules of a sheet: one header record saying how many there are, then one
/// record per rule with its bounds, its messages and the cells it covers.
/// </summary>
internal static class DataValidationWriter
{
    /// <summary>Writes the whole block, or nothing when the sheet has no rules.</summary>
    public static void Write(BiffWriter writer, Worksheet sheet, PtgEncoder encoder, Action<WorkbookWarning> warn)
    {
        if (sheet.Validations.Count == 0)
            return;

        // Each rule is encoded first: one that cannot be written must not be counted in the
        // header, and the header comes before them all.
        var encoded = new List<(DataValidation Rule, byte[]? First, byte[]? Second)>(sheet.Validations.Count);
        foreach (DataValidation rule in sheet.Validations)
        {
            if (rule.Ranges.Count == 0)
                continue;

            byte[]? first = EncodeFormula(rule, rule.Formula1, encoder, warn, sheet);
            if (rule.Formula1 is { Length: > 0 } && first is null)
                continue;

            byte[]? second = EncodeFormula(rule, rule.Formula2, encoder, warn, sheet);
            encoded.Add((rule, first, second));
        }

        if (encoded.Count == 0)
            return;

        WriteHeader(writer, encoded.Count);
        foreach ((DataValidation rule, byte[]? first, byte[]? second) in encoded)
            WriteRule(writer, rule, first, second);
    }

    /// <summary>Writes the DVal record: where the dropdown object is, and how many rules follow.</summary>
    private static void WriteHeader(BiffWriter writer, int count)
    {
        int start = writer.BeginRecord(BiffRecordType.DVal);
        writer.AppendUInt16(0);                 // No prompt box is being shown.
        writer.AppendUInt32(0);                 // Where that box would sit, in pixels.
        writer.AppendUInt32(0);
        writer.AppendUInt32(0xFFFFFFFF);        // Object id of the dropdown: none of ours.
        writer.AppendUInt32((uint)count);
        writer.EndRecord(start);
    }

    private static void WriteRule(BiffWriter writer, DataValidation rule, byte[]? first, byte[]? second)
    {
        uint flags = (uint)MapType(rule.Type);
        flags |= (uint)MapErrorStyle(rule.ErrorStyle) << 4;
        flags |= (uint)MapOperator(rule.Operator) << 20;

        // A list whose source is quoted text is a literal list rather than a reference, and
        // the format wants to be told so.
        if (rule.Type == DataValidationType.List && IsLiteralList(rule.Formula1))
            flags |= 0x00000080;
        if (rule.AllowBlank)
            flags |= 0x00000100;
        if (rule.Type == DataValidationType.List && !rule.ShowDropDown)
            flags |= 0x00000200;
        if (rule.ShowInputMessage)
            flags |= 0x00040000;
        if (rule.ShowErrorMessage)
            flags |= 0x00080000;

        int start = writer.BeginRecord(BiffRecordType.Dv);
        writer.AppendUInt32(flags);
        WriteMessage(writer, rule.PromptTitle);
        WriteMessage(writer, rule.ErrorTitle);
        WriteMessage(writer, rule.Prompt);
        WriteMessage(writer, rule.Error);
        WriteFormula(writer, first);
        WriteFormula(writer, second);

        writer.AppendUInt16(rule.Ranges.Count);
        foreach (RangeRef range in rule.Ranges)
        {
            writer.AppendUInt16(Math.Min(range.Start.Row, 65536) - 1);
            writer.AppendUInt16(Math.Min(range.End.Row, 65536) - 1);
            writer.AppendUInt16(Math.Min(range.Start.Column, 256) - 1);
            writer.AppendUInt16(Math.Min(range.End.Column, 256) - 1);
        }

        writer.EndRecord(start);
    }

    /// <summary>
    /// Encodes one of a rule's formulas. A literal list — <c>"Yes,No"</c> — is not a formula at
    /// all in this format: it is a single string with the choices separated by NUL.
    /// </summary>
    private static byte[]? EncodeFormula(
        DataValidation rule, string? text, PtgEncoder encoder, Action<WorkbookWarning> warn, Worksheet sheet)
    {
        if (string.IsNullOrEmpty(text))
            return null;

        if (rule.Type == DataValidationType.List && IsLiteralList(text))
            return EncodeLiteralList(text);

        byte[]? encoded = encoder.Encode(text, rule.Ranges[0].Start);
        if (encoded is null)
        {
            warn(new WorkbookWarning(
                "UnwritableFormula",
                $"A validation rule on '{sheet.Name}' uses a formula with no encoding in the .xls format; " +
                "the rule was left out."));
        }

        return encoded;
    }

    /// <summary>Whether a list source is a quoted set of choices rather than a reference.</summary>
    private static bool IsLiteralList(string? formula) =>
        formula is { Length: >= 2 } text && text[0] == '"' && text[^1] == '"';

    /// <summary>Packs the choices of a literal list into the one string token the format uses.</summary>
    private static byte[] EncodeLiteralList(string quoted)
    {
        string choices = quoted[1..^1].Replace(',', '\0');
        using var tokens = new BiffWriter(64);
        tokens.AppendByte(0x17);                // PtgStr.
        tokens.AppendByte((byte)Math.Min(choices.Length, 255));
        tokens.AppendByte(BiffWriter.IsCompressible(choices) ? (byte)0x00 : (byte)0x01);
        if (BiffWriter.IsCompressible(choices))
        {
            foreach (char c in choices)
                tokens.AppendByte((byte)c);
        }
        else
        {
            foreach (char c in choices)
                tokens.AppendUInt16(c);
        }

        return tokens.ToArray();
    }

    /// <summary>Writes a formula and its length; an absent one is a length of zero.</summary>
    private static void WriteFormula(BiffWriter writer, byte[]? encoded)
    {
        writer.AppendUInt16(encoded?.Length ?? 0);
        writer.AppendUInt16(0);                 // Reserved.
        if (encoded is not null)
            writer.Append(encoded);
    }

    /// <summary>Writes one message string. A single NUL is how the format spells "none".</summary>
    private static void WriteMessage(BiffWriter writer, string? text) =>
        writer.AppendUnicodeString(string.IsNullOrEmpty(text) ? "\0" : text);

    private static int MapType(DataValidationType type) => type switch
    {
        DataValidationType.Whole => 1,
        DataValidationType.Decimal => 2,
        DataValidationType.List => 3,
        DataValidationType.Date => 4,
        DataValidationType.Time => 5,
        DataValidationType.TextLength => 6,
        DataValidationType.Custom => 7,
        _ => 0,
    };

    private static int MapOperator(DataValidationOperator op) => op switch
    {
        DataValidationOperator.NotBetween => 1,
        DataValidationOperator.Equal => 2,
        DataValidationOperator.NotEqual => 3,
        DataValidationOperator.GreaterThan => 4,
        DataValidationOperator.LessThan => 5,
        DataValidationOperator.GreaterThanOrEqual => 6,
        DataValidationOperator.LessThanOrEqual => 7,
        _ => 0,
    };

    private static int MapErrorStyle(DataValidationErrorStyle style) => style switch
    {
        DataValidationErrorStyle.Warning => 1,
        DataValidationErrorStyle.Information => 2,
        _ => 0,
    };
}
