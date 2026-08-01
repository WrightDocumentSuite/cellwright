using System.Buffers.Binary;

namespace Cellwright.Xls.Biff;

/// <summary>
/// The little numeric encodings BIFF uses to keep a cell down to four bytes, and the packing of
/// a formula's cached result.
/// </summary>
internal static class BiffValues
{
    /// <summary>
    /// Decodes an RK number (MS-XLS 2.5.166): the top 30 bits of a double, or a 30-bit signed
    /// integer, either of them optionally divided by 100. It exists because most numbers in a
    /// sheet are small and round, and this makes them cost four bytes instead of eight.
    /// </summary>
    public static double DecodeRk(uint rk)
    {
        bool divide = (rk & 0x01) != 0;
        bool isInteger = (rk & 0x02) != 0;

        double value;
        if (isInteger)
        {
            // A 30-bit two's-complement integer sitting in the high bits.
            int number = (int)rk >> 2;
            value = number;
        }
        else
        {
            // The 30 bits are the top of an IEEE double; the rest were zero.
            value = BitConverter.Int64BitsToDouble((long)(rk & 0xFFFFFFFC) << 32);
        }

        return divide ? value / 100.0 : value;
    }

    /// <summary>
    /// Encodes a number as RK, or reports that it does not fit. Only exact round trips are
    /// accepted: a lossy RK would silently change the value in the file.
    /// </summary>
    public static bool TryEncodeRk(double value, out uint rk)
    {
        rk = 0;

        // The integer forms carry 30 bits of two's complement.
        const int Min = -(1 << 29);
        const int Max = (1 << 29) - 1;

        // Negative zero has to go the long way round: the integer form would flatten it to zero.
        bool isNegativeZero = value == 0 && double.IsNegative(value);
        if (value == Math.Floor(value) && value is >= Min and <= Max && !isNegativeZero)
        {
            rk = (uint)(((int)value << 2) | 0x02);
            return true;
        }

        double hundredths = value * 100.0;
        if (hundredths == Math.Floor(hundredths) && hundredths is >= Min and <= Max &&
            (int)hundredths / 100.0 == value)
        {
            rk = (uint)(((int)hundredths << 2) | 0x03);
            return true;
        }

        long bits = BitConverter.DoubleToInt64Bits(value);
        if ((bits & 0x00000000FFFFFFFFL) == 0)
        {
            rk = (uint)(bits >> 32) & 0xFFFFFFFC;
            return true;
        }

        long scaled = BitConverter.DoubleToInt64Bits(value * 100.0);
        if ((scaled & 0x00000000FFFFFFFFL) == 0 &&
            BitConverter.Int64BitsToDouble(scaled) / 100.0 == value)
        {
            rk = ((uint)(scaled >> 32) & 0xFFFFFFFC) | 0x01;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Reads the eight bytes a formula caches its last result in (MS-XLS 2.5.133). A number is
    /// stored outright; anything else is tagged in the first byte with 0xFFFF in the last two.
    /// </summary>
    /// <param name="body">The eight bytes.</param>
    /// <param name="value">The cached value, blank for a string whose text follows separately.</param>
    /// <returns>Whether a <c>String</c> record follows carrying the text.</returns>
    public static bool TryReadCachedValue(ReadOnlySpan<byte> body, out CellValue value)
    {
        value = CellValue.Blank;
        if (body.Length < 8)
            return false;

        if (BinaryPrimitives.ReadUInt16LittleEndian(body[6..]) != 0xFFFF)
        {
            value = CellValue.Number(BinaryPrimitives.ReadDoubleLittleEndian(body));
            return false;
        }

        switch (body[0])
        {
            case 0: // The text arrives in the String record that follows.
                return true;
            case 1:
                value = CellValue.Boolean(body[2] != 0);
                return false;
            case 2:
                value = CellValue.Error(DecodeError(body[2]));
                return false;
            default: // 3: an empty string.
                value = CellValue.Text("");
                return false;
        }
    }

    /// <summary>Writes the cached result of a formula into the eight bytes that carry it.</summary>
    public static void WriteCachedValue(Span<byte> destination, in CellValue value, ExcelDateSystem dateSystem)
    {
        switch (value.Kind)
        {
            case CellValueKind.Number or CellValueKind.DateTime or CellValueKind.TimeSpan:
                BinaryPrimitives.WriteDoubleLittleEndian(destination, value.ToSerialNumber(dateSystem));
                return;

            case CellValueKind.Boolean:
                destination.Clear();
                destination[0] = 1;
                destination[2] = value.GetBoolean() ? (byte)1 : (byte)0;
                BinaryPrimitives.WriteUInt16LittleEndian(destination[6..], 0xFFFF);
                return;

            case CellValueKind.Error:
                destination.Clear();
                destination[0] = 2;
                destination[2] = EncodeError(value.GetError());
                BinaryPrimitives.WriteUInt16LittleEndian(destination[6..], 0xFFFF);
                return;

            // An empty result, and a formula that was never calculated, are the same thing to
            // the format. Both MUST take the empty-cell encoding: leaving the leading byte at
            // zero would claim a string result, and a reader would then demand the String
            // record that follows one.
            case CellValueKind.Blank:
            case CellValueKind.Text when value.GetText().Length == 0:
                destination.Clear();
                destination[0] = 3;
                BinaryPrimitives.WriteUInt16LittleEndian(destination[6..], 0xFFFF);
                return;

            default:
                // A non-empty string: the text goes in the String record that follows.
                destination.Clear();
                BinaryPrimitives.WriteUInt16LittleEndian(destination[6..], 0xFFFF);
                return;
        }
    }

    /// <summary>Maps a BIFF error code onto the model's error value.</summary>
    public static ExcelError DecodeError(byte code) => code switch
    {
        0x00 => ExcelError.Null,
        0x07 => ExcelError.DivisionByZero,
        0x0F => ExcelError.Value,
        0x17 => ExcelError.Reference,
        0x1D => ExcelError.Name,
        0x24 => ExcelError.Number,
        _ => ExcelError.NotAvailable,
    };

    /// <summary>Maps the model's error value onto a BIFF error code.</summary>
    public static byte EncodeError(ExcelError error) => error switch
    {
        ExcelError.Null => 0x00,
        ExcelError.DivisionByZero => 0x07,
        ExcelError.Value => 0x0F,
        ExcelError.Reference => 0x17,
        ExcelError.Name => 0x1D,
        ExcelError.Number => 0x24,
        _ => 0x2A,
    };
}
