using Cellwright.Formulas;
using Cellwright.Formulas.Parsing;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Reading;
using Cellwright.Xls.Writing;

namespace Cellwright.Xls.Formulas;

/// <summary>
/// Encodes a formula's text into the token stream BIFF stores. The text is compiled by the
/// library's own parser, which already produces reverse-Polish order — the same order the file
/// format wants — so encoding is a walk over the compiled program.
/// </summary>
/// <remarks>
/// Operand classes (value, reference, array) are written as value class throughout, apart from
/// the arguments of functions that need a reference. Excel recomputes the classes when it next
/// recalculates, and every reader tested tolerates the simplification, but a formula whose
/// meaning depends on the distinction — <c>ROW(A1:A5)</c> as an array formula, say — may be
/// evaluated as a single value until Excel touches it.
/// </remarks>
internal sealed class PtgEncoder(XlsWriteContext context)
{
    private readonly BiffWriter _tokens = new(256);
    private readonly List<byte[]> _arrayConstants = [];

    /// <summary>
    /// Encodes formula text. Returns <see langword="null"/> when the text uses something the
    /// format has no token for, so the caller can fall back to writing the value.
    /// </summary>
    /// <param name="text">Formula text, with or without the leading <c>=</c>.</param>
    /// <param name="anchor">Cell the formula sits in.</param>
    public byte[]? Encode(string text, CellRef anchor)
    {
        FormulaProgram program;
        try
        {
            program = FormulaParser.Parse(text, context.Workbook);
        }
        catch (FormulaParseException)
        {
            return null;
        }

        _tokens.Dispose();
        var tokens = new BiffWriter(Math.Max(64, text.Length * 2));
        _arrayConstants.Clear();

        try
        {
            foreach (Instruction instruction in program.Instructions)
            {
                if (!Emit(tokens, program, instruction, anchor))
                    return null;
            }

            foreach (byte[] constant in _arrayConstants)
                tokens.Append(constant);

            return tokens.Length > BiffReader.MaxRecordSize ? null : tokens.ToArray();
        }
        finally
        {
            tokens.Dispose();
        }
    }

    private bool Emit(BiffWriter tokens, FormulaProgram program, in Instruction instruction, CellRef anchor)
    {
        switch (instruction.Code)
        {
            case OpCode.PushNumber:
                EmitNumber(tokens, program.Numbers[instruction.Operand]);
                return true;

            case OpCode.PushString:
                return EmitString(tokens, program.Strings[instruction.Operand]);

            case OpCode.PushBoolean:
                tokens.AppendByte(0x1D);
                tokens.AppendByte((byte)(instruction.Operand != 0 ? 1 : 0));
                return true;

            case OpCode.PushError:
                tokens.AppendByte(0x1C);
                tokens.AppendByte(BiffValues.EncodeError((ExcelError)instruction.Operand));
                return true;

            case OpCode.PushBlank:
                tokens.AppendByte(0x16);
                return true;

            case OpCode.PushReference:
                return EmitReference(tokens, program.References[instruction.Operand], program, anchor);

            case OpCode.PushName:
                return EmitName(tokens, program.Strings[instruction.Operand]);

            case OpCode.PushArray:
                EmitArray(tokens, program.Arrays[instruction.Operand]);
                return true;

            case OpCode.Add: tokens.AppendByte(0x03); return true;
            case OpCode.Subtract: tokens.AppendByte(0x04); return true;
            case OpCode.Multiply: tokens.AppendByte(0x05); return true;
            case OpCode.Divide: tokens.AppendByte(0x06); return true;
            case OpCode.Power: tokens.AppendByte(0x07); return true;
            case OpCode.Concat: tokens.AppendByte(0x08); return true;
            case OpCode.Less: tokens.AppendByte(0x09); return true;
            case OpCode.LessOrEqual: tokens.AppendByte(0x0A); return true;
            case OpCode.Equal: tokens.AppendByte(0x0B); return true;
            case OpCode.GreaterOrEqual: tokens.AppendByte(0x0C); return true;
            case OpCode.Greater: tokens.AppendByte(0x0D); return true;
            case OpCode.NotEqual: tokens.AppendByte(0x0E); return true;
            case OpCode.Intersect: tokens.AppendByte(0x0F); return true;
            case OpCode.Union: tokens.AppendByte(0x10); return true;
            case OpCode.Range: tokens.AppendByte(0x11); return true;
            // PtgUplus is 0x12, PtgUminus 0x13 (MS-XLS 2.5.198.93 and 2.5.198.95).
            case OpCode.Plus: tokens.AppendByte(0x12); return true;
            case OpCode.Negate: tokens.AppendByte(0x13); return true;
            case OpCode.Percent: tokens.AppendByte(0x14); return true;

            case OpCode.Call:
                return EmitCall(tokens, program.FunctionNames[instruction.Operand], instruction.ArgCount);

            default:
                // LET bindings, structured references and locals have no token in this format.
                return false;
        }
    }

    private static void EmitNumber(BiffWriter tokens, double value)
    {
        // A small whole number gets the two-byte integer token, which is what Excel writes and
        // what keeps a simple formula short.
        if (value is >= 0 and <= 65535 && value == Math.Floor(value))
        {
            tokens.AppendByte(0x1E);
            tokens.AppendUInt16((int)value);
            return;
        }

        tokens.AppendByte(0x1F);
        tokens.AppendDouble(value);
    }

    private static bool EmitString(BiffWriter tokens, string text)
    {
        if (text.Length > 255)
            return false;

        tokens.AppendByte(0x17);
        tokens.AppendShortUnicodeString(text);
        return true;
    }

    private bool EmitReference(BiffWriter tokens, in FormulaRef reference, FormulaProgram program, CellRef anchor)
    {
        // A whole column or row has no token of its own: the format writes it as the widest
        // range that fits, which is exactly how Excel stores A:C.
        (int startRow, int startColumn, int endRow, int endColumn) = Extent(reference);
        bool isRange = reference.IsRange || reference.IsWholeColumn || reference.IsWholeRow;

        int? externSheet = null;
        if (reference.SheetId != FormulaRef.CurrentSheet)
        {
            string sheetName = reference.SheetNameIndex >= 0 && reference.SheetNameIndex < program.Strings.Length
                ? program.Strings[reference.SheetNameIndex]
                : "";
            externSheet = context.ExternSheetFor(sheetName);
            if (externSheet is null)
                return false;
        }

        _ = anchor;
        if (externSheet is { } index)
        {
            tokens.AppendByte(isRange ? (byte)0x3B : (byte)0x3A);
            tokens.AppendUInt16(index);
        }
        else
        {
            tokens.AppendByte(isRange ? (byte)0x25 : (byte)0x24);
        }

        if (isRange)
        {
            // RgceArea puts both rows first and then both columns; the older form interleaves
            // nothing but writes each row in two bytes instead of four (MS-XLSB 2.5.98.89).
            AppendRow(tokens, startRow);
            AppendRow(tokens, endRow);
            tokens.AppendUInt16(EncodeColumn(startColumn, reference.Start, reference));
            tokens.AppendUInt16(EncodeColumn(endColumn, EndPart(reference), reference));
            return true;
        }

        AppendRow(tokens, startRow);
        tokens.AppendUInt16(EncodeColumn(startColumn, reference.Start, reference));
        return true;
    }

    /// <summary>Whether the binary format, whose grid and row field are both wider, is being written.</summary>
    private bool IsBiff12 => context.Version == BiffVersion.Biff12;

    /// <summary>The last row of the target format's grid, which a whole-column reference spans.</summary>
    private int LastRow => IsBiff12 ? 1_048_576 : 65_536;

    /// <summary>The last column, which a whole-row reference spans.</summary>
    private int LastColumn => IsBiff12 ? 16_384 : 256;

    /// <summary>Writes a row number in the width the target format gives it.</summary>
    private void AppendRow(BiffWriter tokens, int row)
    {
        if (IsBiff12)
            tokens.AppendUInt32((uint)(row - 1));
        else
            tokens.AppendUInt16(row - 1);
    }

    /// <summary>Resolves a reference to the four coordinates the format stores.</summary>
    private (int StartRow, int StartColumn, int EndRow, int EndColumn) Extent(in FormulaRef reference)
    {
        if (reference.IsWholeColumn)
        {
            int last = reference.End.Kind == RefPartKind.None ? reference.Start.Column : reference.End.Column;
            return (1, reference.Start.Column, LastRow, last);
        }

        if (reference.IsWholeRow)
        {
            int last = reference.End.Kind == RefPartKind.None ? reference.Start.Row : reference.End.Row;
            return (reference.Start.Row, 1, last, LastColumn);
        }

        return reference.End.Kind == RefPartKind.None
            ? (reference.Start.Row, reference.Start.Column, reference.Start.Row, reference.Start.Column)
            : (reference.Start.Row, reference.Start.Column, reference.End.Row, reference.End.Column);
    }

    private static RefPart EndPart(in FormulaRef reference) =>
        reference.End.Kind == RefPartKind.None ? reference.Start : reference.End;

    /// <summary>
    /// Packs a column number with the two anchor bits. The bits are inverted from what one
    /// might expect: a set bit means the coordinate is relative.
    /// </summary>
    private int EncodeColumn(int column, in RefPart part, in FormulaRef reference)
    {
        // Fourteen bits of column either way, but the older grid only ever fills eight of them.
        int value = Math.Clamp(column - 1, 0, LastColumn - 1);

        // A whole-column reference anchors its rows and a whole-row one anchors its columns:
        // there is nothing for the missing half to be relative to.
        bool rowRelative = !part.RowAbsolute && !reference.IsWholeColumn;
        bool columnRelative = !part.ColumnAbsolute && !reference.IsWholeRow;

        if (rowRelative)
            value |= 0x8000;
        if (columnRelative)
            value |= 0x4000;
        return value;
    }

    private bool EmitName(BiffWriter tokens, string name)
    {
        int? index = context.NameIndexFor(name);
        if (index is null)
            return false;

        tokens.AppendByte(0x23);
        tokens.AppendUInt16(index.Value + 1);
        tokens.AppendUInt16(0);
        return true;
    }

    private void EmitArray(BiffWriter tokens, ArrayLiteral literal)
    {
        tokens.AppendByte(0x20);
        tokens.AppendZeros(7);

        using var constant = new BiffWriter(64);
        constant.AppendByte((byte)Math.Max(0, literal.Columns - 1));
        constant.AppendUInt16(Math.Max(0, literal.Rows - 1));
        foreach (CellValue value in literal.Values)
            AppendArrayElement(constant, value);
        _arrayConstants.Add(constant.ToArray());
    }

    private static void AppendArrayElement(BiffWriter constant, in CellValue value)
    {
        switch (value.Kind)
        {
            case CellValueKind.Text:
                constant.AppendByte(0x02);
                constant.AppendUnicodeString(value.GetText());
                return;

            case CellValueKind.Boolean:
                constant.AppendByte(0x04);
                constant.AppendByte(value.GetBoolean() ? (byte)1 : (byte)0);
                constant.AppendZeros(7);
                return;

            case CellValueKind.Error:
                constant.AppendByte(0x10);
                constant.AppendByte(BiffValues.EncodeError(value.GetError()));
                constant.AppendZeros(7);
                return;

            case CellValueKind.Blank:
                constant.AppendByte(0x00);
                constant.AppendZeros(8);
                return;

            default:
                constant.AppendByte(0x01);
                constant.AppendDouble(value.ToSerialNumber());
                return;
        }
    }

    private static bool EmitCall(BiffWriter tokens, string name, int argumentCount)
    {
        ushort? index = PtgFunctions.GetIndex(name);
        if (index is null)
        {
            // A function this format never had: writing it as a name would change what the
            // formula means, so the cell falls back to its value.
            return false;
        }

        int fixedArity = PtgFunctions.GetArgumentCount(index.Value);
        if (fixedArity >= 0 && fixedArity == argumentCount)
        {
            tokens.AppendByte(0x41);
            tokens.AppendUInt16(index.Value);
            return true;
        }

        if (argumentCount > 255)
            return false;

        tokens.AppendByte(0x42);
        tokens.AppendByte((byte)argumentCount);
        tokens.AppendUInt16(index.Value);
        return true;
    }
}
