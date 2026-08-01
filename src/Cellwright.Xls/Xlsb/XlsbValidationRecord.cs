using System.Buffers.Binary;
using Cellwright.Model;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Xlsb;

/// <summary>
/// Reads the <c>BrtDVal</c> record (MS-XLSB 2.4.356): one input rule and the messages that go
/// with it.
/// </summary>
/// <remarks>
/// The flags word is bit for bit the one the older format uses, so the three enumerations come
/// from <see cref="DataValidationRecord"/> rather than being written out twice. Everything after
/// the flags is laid out differently: the ranges come before the strings instead of after them,
/// the four strings are in a different order, and each is a counted wide string rather than the
/// older format's byte-or-wide pair.
/// </remarks>
internal static class XlsbValidationRecord
{
    public static void Read(
        ReadOnlySpan<byte> body, WorksheetContentSink sink, PtgDecoder decoder, Action<WorkbookWarning> warn)
    {
        if (body.Length < 8)
            return;

        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(body);
        int offset = 4;

        IReadOnlyList<RangeRef> ranges = ReadRanges(body, ref offset);
        if (ranges.Count == 0)
            return;

        string? errorTitle = Biff12Values.ReadNullableString(body, ref offset);
        string? error = Biff12Values.ReadNullableString(body, ref offset);
        string? promptTitle = Biff12Values.ReadNullableString(body, ref offset);
        string? prompt = Biff12Values.ReadNullableString(body, ref offset);

        string? formula1 = ReadFormula(body, ref offset, decoder);
        string? formula2 = ReadFormula(body, ref offset, decoder);

        DataValidationType type = DataValidationRecord.MapType((int)(flags & 0x0F));
        if (type == DataValidationType.List)
            formula1 = DataValidationRecord.UnpackList(formula1);

        var validation = new DataValidation
        {
            Ranges = ranges,
            Type = type,
            Operator = DataValidationRecord.MapOperator((int)((flags >> 20) & 0x0F)),
            Formula1 = formula1,
            Formula2 = formula2,
            AllowBlank = (flags & 0x00000100) != 0,
            ShowDropDown = (flags & 0x00000200) == 0,
            ShowInputMessage = (flags & 0x00040000) != 0,
            ShowErrorMessage = (flags & 0x00080000) != 0,
            ErrorStyle = DataValidationRecord.MapErrorStyle((int)((flags >> 4) & 0x07)),
            PromptTitle = Nullify(promptTitle),
            Prompt = Nullify(prompt),
            ErrorTitle = Nullify(errorTitle),
            Error = Nullify(error),
        };

        if (validation.Type == DataValidationType.Custom && formula1 is null)
        {
            warn(new WorkbookWarning(
                "UndecodedValidation",
                $"A custom validation rule on '{sink.Sheet.Name}' uses a formula this library cannot render and was dropped."));
            return;
        }

        sink.OnDataValidation(validation);
    }

    /// <summary>An <c>UncheckedSqRfX</c> (2.5.156): a count, then that many four-field ranges.</summary>
    private static IReadOnlyList<RangeRef> ReadRanges(ReadOnlySpan<byte> body, ref int offset)
    {
        if (offset + 4 > body.Length)
            return [];

        int count = BinaryPrimitives.ReadInt32LittleEndian(body[offset..]);
        offset += 4;
        if (count is < 0 or > 8192)
            return [];

        var ranges = new List<RangeRef>(Math.Min(count, 64));
        for (int i = 0; i < count && offset + 16 <= body.Length; i++)
        {
            ranges.Add(new RangeRef(
                new CellRef(
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]) + 1,
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 8)..]) + 1),
                new CellRef(
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 4)..]) + 1,
                    (int)BinaryPrimitives.ReadUInt32LittleEndian(body[(offset + 12)..]) + 1)));
            offset += 16;
        }

        return ranges;
    }

    /// <summary>
    /// A <c>DVParsedFormula</c> (2.5.98.8): the tokens, then the ancillary data that some of
    /// them point into. Both have to be stepped over even when the rule has no formula.
    /// </summary>
    private static string? ReadFormula(ReadOnlySpan<byte> body, ref int offset, PtgDecoder decoder)
    {
        if (offset + 4 > body.Length)
            return null;

        int tokens = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
        offset += 4;
        if (tokens < 0 || offset + tokens > body.Length)
            return null;

        int start = offset;
        offset += tokens;

        int extra = 0;
        if (offset + 4 <= body.Length)
        {
            extra = (int)BinaryPrimitives.ReadUInt32LittleEndian(body[offset..]);
            offset += 4;
            extra = extra >= 0 && offset + extra <= body.Length ? extra : 0;
            offset += extra;
        }

        // A rule's formula is anchored at A1: it has no cell of its own to be relative to. The
        // ancillary data follows the tokens, which is where the decoder expects to find it.
        return tokens == 0 ? null : decoder.Decode(body.Slice(start, tokens + extra), new CellRef(1, 1), sheetIndex: 0);
    }

    private static string? Nullify(string? text) => string.IsNullOrEmpty(text) || text == "\0" ? null : text;
}
