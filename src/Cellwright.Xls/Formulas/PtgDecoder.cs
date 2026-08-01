using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Reading;

namespace Cellwright.Xls.Formulas;

/// <summary>
/// Turns the encoded form of a formula back into text. BIFF stores a formula as a parse tree in
/// reverse Polish order — a sequence of "parsed thing" tokens — and this walks it, rebuilding
/// the infix expression a person would recognise.
/// </summary>
/// <remarks>
/// The output has to be text the library's own parser accepts, since that is what recalculates
/// it afterwards. Where the two dialects differ the decoder writes the modern spelling.
/// </remarks>
internal sealed class PtgDecoder(XlsGlobals globals, IReadOnlyList<string> sheetNames)
{
    private readonly List<Fragment> _stack = [];

    /// <summary>
    /// Tables by the number a <c>PtgList</c> names them with, filled as the binary package's
    /// table parts are read. Empty for every other format, which has no such token.
    /// </summary>
    public Dictionary<uint, ExcelTable> Tables { get; } = [];

    /// <summary>A partly-built expression and how tightly it binds, for parenthesising.</summary>
    private readonly record struct Fragment(string Text, int Precedence)
    {
        /// <summary>Wraps the text in brackets when an operator of higher precedence needs it.</summary>
        public string Bracketed(int required) => Precedence < required ? $"({Text})" : Text;
    }

    /// <summary>Precedence of an atom, which never needs brackets.</summary>
    private const int Atomic = 100;

    /// <summary>
    /// Decodes an expression to text. Returns <see langword="null"/> when the token stream holds
    /// something the decoder does not understand, so the caller can fall back to the cached value.
    /// </summary>
    /// <param name="rgce">The token stream.</param>
    /// <param name="anchor">Cell the formula sits in, which relative references are measured from.</param>
    /// <param name="sheetIndex">Zero-based sheet the formula belongs to, for local references.</param>
    public string? Decode(ReadOnlySpan<byte> rgce, CellRef anchor, int sheetIndex)
    {
        _stack.Clear();
        try
        {
            return DecodeCore(rgce, anchor, sheetIndex);
        }
        catch (Exception exception) when (exception is IndexOutOfRangeException or ArgumentOutOfRangeException)
        {
            // A truncated or malformed token stream is a broken file, not a bug worth crashing
            // the load over; the cell keeps its cached value instead.
            return null;
        }
    }

    private string? DecodeCore(ReadOnlySpan<byte> rgce, CellRef anchor, int sheetIndex)
    {
        // Array constants and the like are appended after the token stream, and are consumed in
        // order as the tokens that need them come up.
        int extraOffset = FindExtraDataStart(rgce);
        int extraCursor = extraOffset;

        int offset = 0;
        while (offset < extraOffset)
        {
            byte token = rgce[offset++];

            // Tokens below 0x20 are the operators and literals, identified outright. From 0x20
            // up, the two bits above the low five carry the operand class — value, reference or
            // array — which matters only to an evaluator, so the low five bits say which token
            // this is.
            bool decoded = token < 0x20
                ? StepBase(rgce, ref offset, token)
                : StepClassed(rgce, ref offset, ref extraCursor, (byte)(token & 0x1F), anchor);

            if (!decoded)
                return null;
        }

        _ = sheetIndex;
        return _stack.Count == 1 ? _stack[0].Text : null;
    }

    /// <summary>Handles the tokens below 0x20: operators, literals and the control hints.</summary>
    private bool StepBase(ReadOnlySpan<byte> rgce, ref int offset, byte token)
    {
        switch (token)
        {
            case 0x01: // PtgExp: the cell repeats a shared or array formula stated elsewhere.
                offset += 4;
                return false;

            case 0x02: // PtgTbl: a What-If data table, which has no text form.
                offset += 4;
                return false;

            case 0x03: PushBinary("+", 4); return true;
            case 0x04: PushBinary("-", 4); return true;
            case 0x05: PushBinary("*", 5); return true;
            case 0x06: PushBinary("/", 5); return true;
            case 0x07: PushBinary("^", 6); return true;
            case 0x08: PushBinary("&", 3); return true;
            case 0x09: PushBinary("<", 2); return true;
            case 0x0A: PushBinary("<=", 2); return true;
            case 0x0B: PushBinary("=", 2); return true;
            case 0x0C: PushBinary(">=", 2); return true;
            case 0x0D: PushBinary(">", 2); return true;
            case 0x0E: PushBinary("<>", 2); return true;
            case 0x0F: PushBinary(" ", 8); return true;   // Intersection, written as a space.
            case 0x10: PushBinary(",", 7); return true;   // Union.
            case 0x11: PushBinary(":", 9); return true;   // Range between two references.

            // PtgUplus is 0x12 and PtgUminus 0x13, not the other way round (MS-XLS 2.5.198.93
            // and 2.5.198.95). Swapping them flips the sign of every unary term in the file.
            case 0x12: PushPrefix("+"); return true;
            case 0x13: PushPrefix("-"); return true;
            case 0x14: PushPostfix("%"); return true;

            case 0x15: // PtgParen: the file remembers where a person typed brackets.
                if (_stack.Count == 0)
                    return false;
                _stack[^1] = new Fragment($"({_stack[^1].Text})", Atomic);
                return true;

            case 0x16: // PtgMissArg: an argument left empty, as in IF(A1,,2).
                Push("", Atomic);
                return true;

            case 0x17: // PtgStr.
                Push(QuoteString(ReadInlineString(rgce, ref offset)), Atomic);
                return true;

            // The one byte the two formats spend differently. BIFF8 puts PtgSxName here, which
            // only appears on a macro sheet and has no text form; the binary package puts
            // PtgList, a structured reference. They also decline at different widths, and
            // getting the width wrong would misread every token after it rather than just this
            // one, so the version matters even where the answer is the same.
            case 0x18:
            {
                if (!IsBiff12)
                {
                    offset += 4;
                    return false;
                }

                ReadOnlySpan<byte> reference = rgce.Slice(offset, Math.Min(13, rgce.Length - offset));
                offset += 13;
                if (StructuredReference(reference) is not { } text)
                    return false;

                Push(text, Atomic);
                return true;
            }

            case 0x19: // PtgAttr: control and display hints.
                return ReadAttribute(rgce, ref offset);

            case 0x1C: // PtgErr.
                Push(ErrorText(rgce[offset++]), Atomic);
                return true;

            case 0x1D: // PtgBool.
                Push(rgce[offset++] != 0 ? "TRUE" : "FALSE", Atomic);
                return true;

            case 0x1E: // PtgInt.
                Push(BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]).ToString(CultureInfo.InvariantCulture), Atomic);
                offset += 2;
                return true;

            case 0x1F: // PtgNum.
                Push(FormatNumber(BinaryPrimitives.ReadDoubleLittleEndian(rgce[offset..])), Atomic);
                offset += 8;
                return true;

            default:
                return false;
        }
    }

    /// <summary>
    /// Handles the tokens from 0x20 up, whose id sits in the low five bits: the references, the
    /// names and the function calls.
    /// </summary>
    private bool StepClassed(ReadOnlySpan<byte> rgce, ref int offset, ref int extraCursor, byte id, CellRef anchor)
    {
        switch (id)
        {
            case 0x00: // PtgArray: the values follow the token stream.
                offset += IsBiff12 ? 6 : 7;
                Push(ReadArrayConstant(rgce, ref extraCursor), Atomic);
                return true;

            case 0x01: // PtgFunc: a built-in function with a fixed number of arguments.
            {
                ushort index = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
                offset += 2;
                string? name = PtgFunctions.GetName(index);
                if (name is null)
                    return false;

                int arity = PtgFunctions.GetArgumentCount(index);
                return PushCall(name, arity < 0 ? 0 : arity);
            }

            case 0x02: // PtgFuncVar: the argument count is in the record.
            {
                byte count = rgce[offset];
                ushort index = (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 1)..]) & 0x7FFF);
                offset += 3;
                int arguments = count & 0x7F;

                // Index 255 marks a call to something outside the built-in table: the first
                // argument is the name, pushed by the PtgName token before it.
                if (index == 255)
                    return PushUserCall(arguments);

                string? name = PtgFunctions.GetName(index);
                return name is not null && PushCall(name, arguments);
            }

            case 0x03: // PtgName.
            {
                int nameIndex = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
                offset += globals.Version >= BiffVersion.Biff8 ? 4 : 14;
                Push(NameAt(nameIndex), Atomic);
                return true;
            }

            case 0x04: // PtgRef.
            {
                CellRef reference = ReadRefBiff8(rgce, ref offset, anchor, out bool rowAbsolute, out bool columnAbsolute);
                Push(FormatCell(reference, rowAbsolute, columnAbsolute), Atomic);
                return true;
            }

            case 0x05: // PtgArea.
            {
                string text = ReadAreaBiff8(rgce, ref offset, anchor);
                Push(text, Atomic);
                return true;
            }

            case 0x06: // PtgMemArea: a cached union, whose parts follow as ordinary tokens.
                offset += 6;
                return true;

            case 0x07: // PtgMemErr.
                offset += IsBiff12 ? 4 : 6;
                Push(ExcelError.Reference.ToText(), Atomic);
                return true;

            case 0x08: // PtgMemNoMem.
                offset += 6;
                return true;

            case 0x09: // PtgMemFunc: the expression that follows produces a reference.
                offset += 2;
                return true;

            case 0x0A: // PtgRefErr.
                offset += IsBiff12 ? 6 : 4;
                Push(ExcelError.Reference.ToText(), Atomic);
                return true;

            case 0x0B: // PtgAreaErr.
                offset += IsBiff12 ? 12 : 8;
                Push(ExcelError.Reference.ToText(), Atomic);
                return true;

            case 0x0C: // PtgRefN: relative to the cell, as inside a shared formula.
            {
                CellRef reference = ReadRefBiff8(rgce, ref offset, anchor, out bool rowAbsolute, out bool columnAbsolute, relative: true);
                Push(FormatCell(reference, rowAbsolute, columnAbsolute), Atomic);
                return true;
            }

            case 0x0D: // PtgAreaN.
            {
                string text = ReadAreaBiff8(rgce, ref offset, anchor, relative: true);
                Push(text, Atomic);
                return true;
            }

            case 0x0E: // PtgMemAreaN.
            case 0x0F: // PtgMemNoMemN.
                offset += 2;
                return true;

            case 0x19: // PtgNameX: a name borrowed from another workbook or from an add-in.
            {
                string? name = globals.Version >= BiffVersion.Biff8
                    ? ExternalName(rgce, offset)
                    : null;
                offset += globals.Version >= BiffVersion.Biff8 ? 6 : 24;
                Push(name ?? "#REF!", Atomic);
                return true;
            }

            case 0x1A: // PtgRef3d.
            {
                int externSheet = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
                offset += 2;
                CellRef reference = ReadRefBiff8(rgce, ref offset, anchor, out bool rowAbsolute, out bool columnAbsolute);
                Push(SheetPrefix(externSheet) + FormatCell(reference, rowAbsolute, columnAbsolute), Atomic);
                return true;
            }

            case 0x1B: // PtgArea3d.
            {
                int externSheet = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
                offset += 2;
                string area = ReadAreaBiff8(rgce, ref offset, anchor);
                Push(SheetPrefix(externSheet) + area, Atomic);
                return true;
            }

            case 0x1C: // PtgRefErr3d.
                offset += IsBiff12 ? 8 : 6;
                Push(ExcelError.Reference.ToText(), Atomic);
                return true;

            case 0x1D: // PtgAreaErr3d.
                offset += IsBiff12 ? 14 : 10;
                Push(ExcelError.Reference.ToText(), Atomic);
                return true;

            default:
                return false;
        }
    }

    /// <summary>
    /// Reads a PtgAttr, the token that carries hints rather than content. Only two of them show
    /// up in the text: the shorthand for <c>SUM</c> and the spaces a person typed.
    /// </summary>
    private bool ReadAttribute(ReadOnlySpan<byte> rgce, ref int offset)
    {
        byte flags = rgce[offset++];
        if ((flags & 0x10) != 0) // AttrSum: SUM over one argument, written without the call.
        {
            offset += 2;
            return PushCall("SUM", 1);
        }

        if ((flags & 0x40) != 0) // AttrSpace: whitespace to reproduce.
        {
            byte kind = rgce[offset];
            byte count = rgce[offset + 1];
            offset += 2;
            ApplySpacing(kind, count);
            return true;
        }

        if ((flags & 0x04) != 0) // AttrChoose: a jump table whose length is in the record.
        {
            int count = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
            offset += 2 + ((count + 1) * 2);
            return true;
        }

        // AttrIf, AttrGoto, AttrSkip and the rest: two bytes of jump distance, no text.
        offset += 2;
        return true;
    }

    /// <summary>Reattaches the whitespace a person typed, which Excel preserves verbatim.</summary>
    private void ApplySpacing(byte kind, byte count)
    {
        if (count == 0 || _stack.Count == 0)
            return;

        string padding = new(kind is 2 or 4 or 6 ? '\n' : ' ', count);
        Fragment top = _stack[^1];

        // Odd kinds pad before the token, even ones after the opening bracket of a call.
        _stack[^1] = kind is 0 or 2 or 4
            ? new Fragment(padding + top.Text, top.Precedence)
            : new Fragment(top.Text + padding, top.Precedence);
    }

    private bool PushCall(string name, int arguments)
    {
        if (_stack.Count < arguments)
            return false;

        var builder = new StringBuilder(name).Append('(');
        int start = _stack.Count - arguments;
        for (int i = 0; i < arguments; i++)
        {
            if (i > 0)
                builder.Append(',');
            builder.Append(_stack[start + i].Text);
        }

        _stack.RemoveRange(start, arguments);
        Push(builder.Append(')').ToString(), Atomic);
        return true;
    }

    /// <summary>
    /// Builds a call to something outside the built-in table. The name was pushed first, so it
    /// is the bottom of the argument group rather than an argument.
    /// </summary>
    private bool PushUserCall(int arguments)
    {
        if (arguments < 1 || _stack.Count < arguments)
            return false;

        int start = _stack.Count - arguments;
        string name = _stack[start].Text;
        var builder = new StringBuilder(name).Append('(');
        for (int i = 1; i < arguments; i++)
        {
            if (i > 1)
                builder.Append(',');
            builder.Append(_stack[start + i].Text);
        }

        _stack.RemoveRange(start, arguments);
        Push(builder.Append(')').ToString(), Atomic);
        return true;
    }

    private void PushBinary(string op, int precedence)
    {
        if (_stack.Count < 2)
            return;

        Fragment right = _stack[^1];
        Fragment left = _stack[^2];
        _stack.RemoveRange(_stack.Count - 2, 2);

        // The right operand of a left-associative operator needs brackets at equal precedence:
        // a-(b-c) is not a-b-c.
        string text = op == ":" || op == "," || op == " "
            ? left.Text + op + right.Text
            : left.Bracketed(precedence) + op + right.Bracketed(precedence + 1);
        Push(text, precedence);
    }

    private void PushPrefix(string op)
    {
        if (_stack.Count == 0)
            return;

        Fragment operand = _stack[^1];
        _stack[^1] = new Fragment(op + operand.Bracketed(6), 6);
    }

    private void PushPostfix(string op)
    {
        if (_stack.Count == 0)
            return;

        Fragment operand = _stack[^1];
        _stack[^1] = new Fragment(operand.Bracketed(7) + op, 7);
    }

    private void Push(string text, int precedence) => _stack.Add(new Fragment(text, precedence));

    /// <summary>
    /// Reads the cell reference of a BIFF8 token: a row and a column with two flag bits saying
    /// which halves are anchored.
    /// </summary>
    private CellRef ReadRefBiff8(
        ReadOnlySpan<byte> rgce, ref int offset, CellRef anchor,
        out bool rowAbsolute, out bool columnAbsolute, bool relative = false)
    {
        if (IsBiff12)
            return ReadRefBiff12(rgce, ref offset, anchor, out rowAbsolute, out columnAbsolute, relative);

        int row = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
        ushort column = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 2)..]);
        offset += 4;

        rowAbsolute = (column & 0x8000) == 0;
        columnAbsolute = (column & 0x4000) == 0;
        int columnNumber = column & 0x3FFF;

        // A shared formula stores its references as offsets from the cell it sits in, and the
        // offsets wrap: a column of -1 arrives as 255 in the low byte.
        if (relative)
        {
            if (!rowAbsolute)
                row = WrapRow(anchor.Row - 1 + (short)(ushort)row);
            if (!columnAbsolute)
                columnNumber = WrapColumn(anchor.Column - 1 + (sbyte)columnNumber);
        }

        return new CellRef(
            Math.Clamp(row + 1, 1, SheetLimits.MaxRow),
            Math.Clamp(columnNumber + 1, 1, SheetLimits.MaxColumn));
    }

    /// <summary>
    /// Reads an <c>RgceLoc</c> (MS-XLSB 2.5.98.91): four bytes of row, then two of column with
    /// the same two flag bits. The wider row is the whole difference, and it is what the modern
    /// grid of a million rows costs.
    /// </summary>
    private static CellRef ReadRefBiff12(
        ReadOnlySpan<byte> rgce, ref int offset, CellRef anchor,
        out bool rowAbsolute, out bool columnAbsolute, bool relative)
    {
        int row = (int)BinaryPrimitives.ReadUInt32LittleEndian(rgce[offset..]);
        ushort column = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 4)..]);
        offset += 6;

        rowAbsolute = (column & 0x8000) == 0;
        columnAbsolute = (column & 0x4000) == 0;
        int columnNumber = column & 0x3FFF;

        if (relative)
        {
            if (!rowAbsolute)
                row = WrapRowBiff12(anchor.Row - 1 + row);
            if (!columnAbsolute)
                columnNumber = WrapColumnBiff12(anchor.Column - 1 + (short)(ushort)columnNumber);
        }

        return new CellRef(
            Math.Clamp(row + 1, 1, SheetLimits.MaxRow),
            Math.Clamp(columnNumber + 1, 1, SheetLimits.MaxColumn));
    }

    private string ReadAreaBiff8(ReadOnlySpan<byte> rgce, ref int offset, CellRef anchor, bool relative = false)
    {
        int firstRow;
        int lastRow;
        ushort firstColumn;
        ushort lastColumn;
        if (IsBiff12)
        {
            // RgceArea (MS-XLSB 2.5.98.89): both rows first, four bytes each, then both columns.
            firstRow = (int)BinaryPrimitives.ReadUInt32LittleEndian(rgce[offset..]);
            lastRow = (int)BinaryPrimitives.ReadUInt32LittleEndian(rgce[(offset + 4)..]);
            firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 8)..]);
            lastColumn = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 10)..]);
            offset += 12;
        }
        else
        {
            firstRow = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
            lastRow = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 2)..]);
            firstColumn = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 4)..]);
            lastColumn = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 6)..]);
            offset += 8;
        }

        bool firstRowAbsolute = (firstColumn & 0x8000) == 0;
        bool firstColumnAbsolute = (firstColumn & 0x4000) == 0;
        bool lastRowAbsolute = (lastColumn & 0x8000) == 0;
        bool lastColumnAbsolute = (lastColumn & 0x4000) == 0;
        int firstColumnNumber = firstColumn & 0x3FFF;
        int lastColumnNumber = lastColumn & 0x3FFF;

        if (relative && IsBiff12)
        {
            if (!firstRowAbsolute)
                firstRow = WrapRowBiff12(anchor.Row - 1 + firstRow);
            if (!lastRowAbsolute)
                lastRow = WrapRowBiff12(anchor.Row - 1 + lastRow);
            if (!firstColumnAbsolute)
                firstColumnNumber = WrapColumnBiff12(anchor.Column - 1 + (short)(ushort)firstColumnNumber);
            if (!lastColumnAbsolute)
                lastColumnNumber = WrapColumnBiff12(anchor.Column - 1 + (short)(ushort)lastColumnNumber);
        }
        else if (relative)
        {
            if (!firstRowAbsolute)
                firstRow = WrapRow(anchor.Row - 1 + (short)(ushort)firstRow);
            if (!lastRowAbsolute)
                lastRow = WrapRow(anchor.Row - 1 + (short)(ushort)lastRow);
            if (!firstColumnAbsolute)
                firstColumnNumber = WrapColumn(anchor.Column - 1 + (sbyte)firstColumnNumber);
            if (!lastColumnAbsolute)
                lastColumnNumber = WrapColumn(anchor.Column - 1 + (sbyte)lastColumnNumber);
        }

        var start = new CellRef(
            Math.Clamp(firstRow + 1, 1, SheetLimits.MaxRow),
            Math.Clamp(firstColumnNumber + 1, 1, SheetLimits.MaxColumn));
        var end = new CellRef(
            Math.Clamp(lastRow + 1, 1, SheetLimits.MaxRow),
            Math.Clamp(lastColumnNumber + 1, 1, SheetLimits.MaxColumn));

        // A whole column or row is stored as an area spanning the grid, and that is the only
        // way either binary format can say it. Spelling it back out as A$1:A$65536 would name
        // the same cells, but it stops both Excel and this library's engine from recognising
        // the shape and clipping it to the rows actually in use.
        int rows = IsBiff12 ? Biff12Rows : BiffRows;
        int columns = IsBiff12 ? Biff12Columns : BiffColumns;
        if (firstRow == 0 && lastRow == rows - 1)
            return $"{ColumnPart(start, firstColumnAbsolute)}:{ColumnPart(end, lastColumnAbsolute)}";
        if (firstColumnNumber == 0 && lastColumnNumber == columns - 1)
            return $"{RowPart(start, firstRowAbsolute)}:{RowPart(end, lastRowAbsolute)}";

        return FormatCell(start, firstRowAbsolute, firstColumnAbsolute) + ":" +
               FormatCell(end, lastRowAbsolute, lastColumnAbsolute);
    }

    /// <summary>
    /// Renders a <c>PtgList</c> (MS-XLSB 2.5.98.52) as the <c>Table[Column]</c> a person wrote,
    /// or <see langword="null"/> when the table it names is not one this file defines.
    /// </summary>
    /// <remarks>
    /// The token carries numbers, not names: the table by its identifier and the columns by
    /// their position within it. Everything that makes the text readable comes from the table
    /// parts, which is why a workbook whose tables were not read can only decline.
    /// </remarks>
    private string? StructuredReference(ReadOnlySpan<byte> token)
    {
        if (token.Length < 13)
            return null;

        // Past the eptg byte and the supporting-link index, which names the sheet the table is
        // on — and a reference into another workbook's table is one this reader cannot follow.
        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(token[3..]);
        bool invalid = (flags & 0x0400) != 0;
        bool nonresident = (flags & 0x0800) != 0;
        if (invalid || nonresident)
            return null;

        uint id = BinaryPrimitives.ReadUInt32LittleEndian(token[5..]);
        if (!Tables.TryGetValue(id, out ExcelTable? table))
            return null;

        int columns = flags & 0x0003;
        int rowType = (flags >> 2) & 0x1F;
        int first = BinaryPrimitives.ReadUInt16LittleEndian(token[9..]);
        int last = BinaryPrimitives.ReadUInt16LittleEndian(token[11..]);

        // A reference is made of parts: which rows of the table, then which columns. Excel
        // spells a lone part bare and two or more each in brackets of their own, so
        // Sales[Amount] and Sales[[#Headers],[Amount]] are one shape with a different count.
        List<string> parts = [.. Regions(rowType)];
        if (Span(table, columns, first, last) is { } span)
            parts.Add(span);

        if (parts.Count == 0)
            return table.Name;

        // A pair of columns is already bracketed twice and is never bare, even on its own.
        bool bare = parts.Count == 1 && !parts[0].StartsWith('[');
        string inner = bare
            ? parts[0]
            : string.Join(",", parts.Select(static part => part.StartsWith('[') ? part : $"[{part}]"));

        return $"{table.Name}[{inner}]";
    }

    /// <summary>
    /// The rows a reference covers (<c>PtgRowType</c>, MS-XLSB 2.5.98.73). Two of the values
    /// name two regions at once, which is why this yields a sequence rather than one name.
    /// </summary>
    private static string[] Regions(int rowType) => rowType switch
    {
        0x01 => ["#All"],
        0x02 => ["#Headers"],
        0x06 => ["#Headers", "#Data"],
        0x08 => ["#Totals"],
        0x0C => ["#Data", "#Totals"],
        0x10 => ["#This Row"],
        _ => [],
    };

    /// <summary>The columns a reference covers, or null when it covers all of them.</summary>
    private static string? Span(ExcelTable table, int columns, int first, int last)
    {
        if (columns == 0 || Name(table, first) is not { } start)
            return null;
        if (columns == 1 || Name(table, last) is not { } end || end == start)
            return start;

        return $"[{start}]:[{end}]";
    }

    private static string? Name(ExcelTable table, int index) =>
        (uint)index < (uint)table.Columns.Count ? table.Columns[index] : null;

    private const int BiffRows = 65536;
    private const int BiffColumns = 256;

    /// <summary>The modern grid, which the binary package format writes against.</summary>
    private const int Biff12Rows = 1_048_576;
    private const int Biff12Columns = 16_384;

    private static int WrapRowBiff12(int row) => ((row % Biff12Rows) + Biff12Rows) % Biff12Rows;

    private static int WrapColumnBiff12(int column) =>
        ((column % Biff12Columns) + Biff12Columns) % Biff12Columns;

    /// <summary>
    /// Brings a relative reference that ran off an edge back onto the sheet. Excel wraps rather
    /// than clamps: a formula one column left of column A reads the last column, and copying a
    /// block of formulas near an edge relies on it.
    /// </summary>
    private static int WrapRow(int row) => ((row % BiffRows) + BiffRows) % BiffRows;

    private static int WrapColumn(int column) => ((column % BiffColumns) + BiffColumns) % BiffColumns;

    private static string ColumnPart(CellRef cell, bool absolute) =>
        (absolute ? "$" : "") + ColumnName.Format(cell.Column);

    private static string RowPart(CellRef cell, bool absolute) =>
        (absolute ? "$" : "") + cell.Row.ToString(CultureInfo.InvariantCulture);

    private static string FormatCell(CellRef cell, bool rowAbsolute, bool columnAbsolute)
    {
        var builder = new StringBuilder(8);
        if (columnAbsolute)
            builder.Append('$');
        builder.Append(ColumnName.Format(cell.Column));
        if (rowAbsolute)
            builder.Append('$');
        builder.Append(cell.Row.ToString(CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    /// <summary>
    /// Resolves a PtgNameX to the name it stands for. The token holds an ExternSheet index and
    /// a one-based position in the ExternName records of the workbook that entry belongs to;
    /// for the add-in book those records hold function names, so this is what turns a call back
    /// into <c>EOMONTH(…)</c> rather than a reference that no longer exists.
    /// </summary>
    private string? ExternalName(ReadOnlySpan<byte> rgce, int offset)
    {
        if (offset + 4 > rgce.Length)
            return null;

        int externSheetIndex = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
        int nameIndex = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 2)..]);
        if ((uint)externSheetIndex >= (uint)globals.ExternSheets.Count)
            return null;

        int supBook = globals.ExternSheets[externSheetIndex].SupBook;
        if ((uint)supBook >= (uint)globals.SupBooks.Count)
            return null;

        // Only the add-in book's names are function names that mean the same thing here. A name
        // defined in another workbook is a reference this library cannot follow, and Excel
        // itself caches #REF! for one whose file has gone, so that is what it becomes.
        if (!globals.SupBooks[supBook].IsAddin)
            return null;

        List<string> names = globals.SupBooks[supBook].Names;
        if (nameIndex < 1 || nameIndex > names.Count)
            return null;

        string name = names[nameIndex - 1];
        return name.Length == 0 ? null : name;
    }

    /// <summary>
    /// Builds the <c>Sheet1!</c> part of a 3-D reference from an ExternSheet index. A reference
    /// spanning several tabs prints as a range of names, which is how Excel writes it too, and
    /// one into another workbook carries that workbook in brackets ahead of the tab.
    /// </summary>
    private string SheetPrefix(int externSheetIndex)
    {
        if ((uint)externSheetIndex >= (uint)globals.ExternSheets.Count)
            return "";

        (int supBook, int first, int last) = globals.ExternSheets[externSheetIndex];
        if (first < 0)
            return "#REF!";

        // Sheet indexes of an entry belonging to another workbook count that workbook's tabs.
        // Resolving them against ours would name a real local sheet and read real local data,
        // so the reference has to be built from the other book's list or not at all.
        IReadOnlyList<string> tabs = sheetNames;
        string book = "";
        if ((uint)supBook < (uint)globals.SupBooks.Count && globals.SupBooks[supBook].IsExternal)
        {
            XlsSupBook other = globals.SupBooks[supBook];
            tabs = other.SheetNames;
            book = $"[{other.Path}]";
        }

        string firstName = SheetNameAt(tabs, first);
        string text = last == first || last < 0 ? firstName : $"{firstName}:{SheetNameAt(tabs, last)}";
        text = book + text;
        return NeedsQuoting(text) ? $"'{text.Replace("'", "''", StringComparison.Ordinal)}'!" : text + "!";
    }

    private static string SheetNameAt(IReadOnlyList<string> tabs, int index) =>
        (uint)index < (uint)tabs.Count ? tabs[index] : "#REF";

    /// <summary>Whether a sheet name has to be quoted to survive being pasted into a formula.</summary>
    private static bool NeedsQuoting(string name)
    {
        if (name.Length == 0)
            return true;
        if (char.IsAsciiDigit(name[0]))
            return true;

        foreach (char c in name)
        {
            if (!char.IsAsciiLetterOrDigit(c) && c != '_' && c != '.')
                return true;
        }

        return false;
    }

    private string NameAt(int oneBasedIndex)
    {
        int index = oneBasedIndex - 1;
        return (uint)index < (uint)globals.Names.Count ? globals.Names[index].Name : "#NAME?";
    }

    private string ReadInlineString(ReadOnlySpan<byte> rgce, ref int offset)
    {
        // The two formats count the characters differently. BIFF8 spends one byte on the count
        // and one on a flag saying whether the characters are one byte or two (MS-XLS 2.5.240);
        // BIFF12 spends two on the count and has no flag, because they are always two bytes
        // (MS-XLSB 2.5.98.74). Reading one as the other leaves the cursor mid-token, so every
        // token after a string in the formula is misread rather than just the string.
        int characters;
        bool wide;
        if (IsBiff12)
        {
            characters = BinaryPrimitives.ReadUInt16LittleEndian(rgce[offset..]);
            offset += 2;
            wide = true;
        }
        else
        {
            characters = rgce[offset++];
            wide = (rgce[offset++] & 0x01) != 0;
        }

        int byteCount = Math.Min(wide ? characters * 2 : characters, rgce.Length - offset);
        string text = wide
            ? Encoding.Unicode.GetString(rgce.Slice(offset, byteCount))
            : Encoding.Latin1.GetString(rgce.Slice(offset, byteCount));
        offset += byteCount;
        return text;
    }

    /// <summary>
    /// Reads an array constant from the data that trails the token stream, and renders it in
    /// the <c>{1,2;3,4}</c> form.
    /// </summary>
    private static string ReadArrayConstant(ReadOnlySpan<byte> rgce, ref int cursor)
    {
        if (cursor + 2 > rgce.Length)
            return "{}";

        int columns = rgce[cursor] + 1;
        int rows = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(cursor + 1)..]) + 1;
        cursor += 3;

        var builder = new StringBuilder("{");
        for (int row = 0; row < rows; row++)
        {
            if (row > 0)
                builder.Append(';');
            for (int column = 0; column < columns; column++)
            {
                if (column > 0)
                    builder.Append(',');
                builder.Append(ReadArrayElement(rgce, ref cursor));
            }
        }

        return builder.Append('}').ToString();
    }

    private static string ReadArrayElement(ReadOnlySpan<byte> rgce, ref int cursor)
    {
        if (cursor >= rgce.Length)
            return "";

        byte kind = rgce[cursor++];
        switch (kind)
        {
            case 0x01 when cursor + 8 <= rgce.Length:
                double number = BinaryPrimitives.ReadDoubleLittleEndian(rgce[cursor..]);
                cursor += 8;
                return FormatNumber(number);

            case 0x02:
                int characters = cursor + 2 <= rgce.Length ? BinaryPrimitives.ReadUInt16LittleEndian(rgce[cursor..]) : 0;
                cursor += 2;
                bool wide = cursor < rgce.Length && (rgce[cursor++] & 0x01) != 0;
                int byteCount = Math.Min(wide ? characters * 2 : characters, Math.Max(0, rgce.Length - cursor));
                string text = wide
                    ? Encoding.Unicode.GetString(rgce.Slice(cursor, byteCount))
                    : Encoding.Latin1.GetString(rgce.Slice(cursor, byteCount));
                cursor += byteCount;
                return QuoteString(text);

            case 0x04:
                bool value = cursor < rgce.Length && rgce[cursor] != 0;
                cursor += 8;
                return value ? "TRUE" : "FALSE";

            case 0x10:
                string error = cursor < rgce.Length ? ErrorText(rgce[cursor]) : "#N/A";
                cursor += 8;
                return error;

            default:
                cursor += 8;
                return "";
        }
    }

    /// <summary>
    /// Finds where the token stream ends and the array constants begin. Only PtgArray tokens
    /// put data there, so the stream is walked once to size it.
    /// </summary>
    private int FindExtraDataStart(ReadOnlySpan<byte> rgce)
    {
        int offset = 0;
        while (offset < rgce.Length)
        {
            byte token = rgce[offset];
            int size = TokenSize(rgce, offset);
            if (size <= 0)
                return rgce.Length;

            offset += size;
            _ = token;
        }

        return Math.Min(offset, rgce.Length);
    }

    /// <summary>Length of one token including its id byte, or -1 when the token is unknown.</summary>
    private int TokenSize(ReadOnlySpan<byte> rgce, int offset)
    {
        byte token = rgce[offset];
        byte id = (byte)(token & 0x1F);
        bool classed = token >= 0x20;

        if (!classed)
        {
            switch (token)
            {
                case >= 0x03 and <= 0x16: return 1;
                case 0x17: // PtgStr, counted differently by the two formats. See ReadInlineString.
                    if (offset + 3 > rgce.Length)
                        return -1;
                    return IsBiff12
                        ? 3 + (2 * BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 1)..]))
                        : 3 + (((rgce[offset + 2] & 0x01) != 0 ? 2 : 1) * rgce[offset + 1]);
                case 0x19: return AttributeSize(rgce, offset);
                case 0x1C or 0x1D: return 2;
                case 0x1E: return 3;
                case 0x1F: return 9;
                case 0x01 or 0x02: return 5;
                case 0x18: return IsBiff12 ? 14 : 5;         // PtgList, or PtgSxName before it.
                default: return -1;
            }
        }

        if (IsBiff12)
            return Biff12TokenSize(id);

        return id switch
        {
            0x00 => 8,                                                       // PtgArray.
            0x01 => 3,                                                       // PtgFunc.
            0x02 => 4,                                                       // PtgFuncVar.
            0x03 => globals.Version >= BiffVersion.Biff8 ? 5 : 15,           // PtgName.
            0x04 or 0x0A or 0x0C => 5,                                       // PtgRef, RefErr, RefN.
            0x05 or 0x0B or 0x0D => 9,                                       // PtgArea, AreaErr, AreaN.
            0x06 or 0x07 or 0x08 => 7,                                       // PtgMemArea and friends.
            0x09 or 0x0E or 0x0F => 3,                                       // PtgMemFunc and the N forms.
            0x19 => globals.Version >= BiffVersion.Biff8 ? 7 : 25,           // PtgNameX.
            0x1A or 0x1C => 7,                                               // PtgRef3d, PtgRefErr3d.
            0x1B or 0x1D => 11,                                              // PtgArea3d, PtgAreaErr3d.
            _ => -1,
        };
    }

    /// <summary>
    /// Sizes of the classed tokens in the binary package format. Most match BIFF8, but every
    /// token carrying a row is four bytes wider, and several of the ones that carry nothing at
    /// all were re-padded, so the whole table is spelled out rather than patched.
    /// </summary>
    private static int Biff12TokenSize(byte id) => id switch
    {
        0x00 => 7,                              // PtgArray: four unused bytes and two more.
        0x01 => 3,                              // PtgFunc.
        0x02 => 4,                              // PtgFuncVar.
        0x03 => 5,                              // PtgName.
        0x04 or 0x0A or 0x0C => 7,              // PtgRef, PtgRefErr, PtgRefN: RgceLoc is six bytes.
        0x05 or 0x0B or 0x0D => 13,             // PtgArea, PtgAreaErr, PtgAreaN: RgceArea is twelve.
        0x06 or 0x08 => 7,                      // PtgMemArea, PtgMemNoMem: unused and a count.
        0x07 => 5,                              // PtgMemErr: an error code and padding.
        0x09 or 0x0E or 0x0F => 3,              // PtgMemFunc and the N forms.
        0x19 => 7,                              // PtgNameX.
        0x1A => 9,                              // PtgRef3d.
        0x1B => 15,                             // PtgArea3d.
        0x1C => 9,                              // PtgRefErr3d.
        0x1D => 15,                             // PtgAreaErr3d.
        _ => -1,
    };

    /// <summary>Whether this workbook's tokens are the ones a binary package writes.</summary>
    private bool IsBiff12 => globals.Version == BiffVersion.Biff12;

    private static int AttributeSize(ReadOnlySpan<byte> rgce, int offset)
    {
        if (offset + 2 > rgce.Length)
            return -1;

        byte flags = rgce[offset + 1];
        if ((flags & 0x04) == 0)
            return 4;

        // AttrChoose carries a jump table sized by the count that follows.
        if (offset + 4 > rgce.Length)
            return -1;
        int count = BinaryPrimitives.ReadUInt16LittleEndian(rgce[(offset + 2)..]);
        return 4 + ((count + 1) * 2);
    }

    private static string QuoteString(string text) =>
        "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";

    private static string FormatNumber(double value) =>
        value.ToString("R", CultureInfo.InvariantCulture);

    private static string ErrorText(byte code) => BiffValues.DecodeError(code).ToText();
}
