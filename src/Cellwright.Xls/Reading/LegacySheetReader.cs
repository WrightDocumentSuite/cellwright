using System.Buffers.Binary;
using System.Text;
using Cellwright.Model;
using Cellwright.Styles;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Reads a workbook from Excel 2.x through 4.0. These files have no container and no global
/// substream: the whole workbook is one record stream holding one sheet, with the style records
/// interleaved among the cells.
/// </summary>
/// <remarks>
/// Values and the number formats that make sense of them are read; formulas are not. Their
/// token stream uses an older encoding that would need a decoder of its own, and the cached
/// result is already in the file, so each formula cell keeps the value it last showed.
/// </remarks>
internal static class LegacySheetReader
{
    /// <summary>Reads the single sheet of a pre-BIFF5 file into the workbook.</summary>
    public static void Read(
        Workbook workbook, BiffReader reader, BiffVersion version, XlsLoadOptions options, Action<WorkbookWarning> warn)
    {
        string generation = version switch
        {
            BiffVersion.Biff2 => "2.x",
            BiffVersion.Biff3 => "3.0",
            _ => "4.0",
        };

        warn(new WorkbookWarning(
            "LegacyBiff",
            $"The workbook was written by Excel {generation}. Its values and number formats were read; formulas " +
            "kept their last computed value instead of their expression."));

        Worksheet sheet = WorkbookLoader.AddSheetSafe(workbook, "Sheet1", warn);
        var globals = new XlsGlobals
        {
            Version = version,
            Encoding = ResolveFallback(options.FallbackCodePage),
        };

        var sink = new WorksheetContentSink(sheet, styleMap: [], relationships: [], preserveMarkup: false, warn);
        var formats = new Dictionary<int, string>();
        var fonts = new List<XlsFontRecord>();
        var cellFormats = new List<XlsXfRecord>();

        // BIFF2 puts a cell's format on the cell itself unless an IXFE record precedes it.
        int pendingXf = -1;

        while (reader.Read())
        {
            ReadOnlySpan<byte> body = reader.Record.Span;
            switch (reader.RecordType)
            {
                case BiffRecordType.Eof:
                    return;

                case BiffRecordType.CodePage:
                    if (body.Length >= 2)
                        globals.Encoding = ResolveCodePage(BinaryPrimitives.ReadUInt16LittleEndian(body), globals.Encoding);
                    break;

                case BiffRecordType.Date1904:
                    if (body.Length >= 2 && BinaryPrimitives.ReadUInt16LittleEndian(body) != 0)
                        workbook.SetDateSystem(ExcelDateSystem.Date1904);
                    break;

                case BiffRecordType.OldFormat:
                    ReadFormat(formats, body, globals.Encoding);
                    break;

                case BiffRecordType.Biff2Xf or BiffRecordType.Biff3Xf or BiffRecordType.Biff4Xf:
                    cellFormats.Add(XlsXfRecord.ParseLegacy(body, version));
                    break;

                case BiffRecordType.Ixfe:
                    pendingXf = body.Length >= 2 ? BinaryPrimitives.ReadUInt16LittleEndian(body) : -1;
                    continue;

                case BiffRecordType.Biff2Integer:
                    ReadInteger(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Biff2Number:
                    ReadNumber(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Biff2Label:
                    ReadLabel(sink, body, ref pendingXf, globals.Encoding, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Biff2BoolErr:
                    ReadBoolErr(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Biff2Formula:
                    ReadFormulaValue(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Rk:
                    ReadRk(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Number:
                    ReadNumber(sink, body, ref pendingXf, workbook, formats, cellFormats);
                    break;

                case BiffRecordType.Label:
                    ReadLabelBiff5(sink, body, ref pendingXf, globals.Encoding, workbook, formats, cellFormats);
                    break;

                default:
                    break;
            }

            pendingXf = -1;
            _ = fonts;
        }
    }

    /// <summary>
    /// Resolves the style of a cell, interning it on first use. These files carry so few
    /// distinct formats that building them lazily costs nothing.
    /// </summary>
    private static (StyleId Style, TemporalKind Temporal) ResolveStyle(
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats, int xfIndex)
    {
        int formatIndex = (uint)xfIndex < (uint)cellFormats.Count ? cellFormats[xfIndex].FormatIndex : 0;
        string code = formats.TryGetValue(formatIndex, out string? custom)
            ? custom
            : BuiltinNumberFormats.GetCode(formatIndex) ?? "";

        var style = new CellStyle { NumberFormat = code };
        return (
            new StyleId(workbook.StylePool.Intern(style)),
            NumberFormatClassifier.Classify(formatIndex, code.Length == 0 ? null : code));
    }

    private static void ReadFormat(Dictionary<int, string> formats, ReadOnlySpan<byte> body, Encoding encoding)
    {
        int cursor = 0;
        formats[formats.Count] = BiffStrings.ReadByteString(body, ref cursor, encoding, lengthBytes: 1);
    }

    /// <summary>
    /// Reads the row, column and format index that open a BIFF2 cell record. The format is
    /// three bytes of cell attributes unless an IXFE record overrode it.
    /// </summary>
    private static (CellRef Cell, int Xf) ReadHeader(ReadOnlySpan<byte> body, int pendingXf)
    {
        var cell = new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
            BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1);
        int xf = pendingXf >= 0 ? pendingXf : body[4] & 0x3F;
        return (cell, xf);
    }

    private static void ReadInteger(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 9)
            return;

        (CellRef cell, int xf) = ReadHeader(body, pendingXf);
        Set(sink, workbook, formats, cellFormats, cell, xf, BinaryPrimitives.ReadUInt16LittleEndian(body[7..]));
    }

    private static void ReadNumber(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 15)
            return;

        (CellRef cell, int xf) = ReadHeader(body, pendingXf);
        Set(sink, workbook, formats, cellFormats, cell, xf, BinaryPrimitives.ReadDoubleLittleEndian(body[7..]));
    }

    private static void ReadRk(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 10)
            return;

        var cell = new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
            BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1);
        int xf = pendingXf >= 0 ? pendingXf : BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        Set(sink, workbook, formats, cellFormats, cell, xf,
            BiffValues.DecodeRk(BinaryPrimitives.ReadUInt32LittleEndian(body[6..])));
    }

    private static void ReadLabel(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf, Encoding encoding,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 8)
            return;

        (CellRef cell, int xf) = ReadHeader(body, pendingXf);
        int cursor = 7;
        string text = BiffStrings.ReadByteString(body, ref cursor, encoding, lengthBytes: 1);
        (StyleId style, _) = ResolveStyle(workbook, formats, cellFormats, xf);
        sink.OnPooledCell(cell, CellValue.Text(text), style, formula: null);
    }

    private static void ReadLabelBiff5(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf, Encoding encoding,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 8)
            return;

        var cell = new CellRef(
            BinaryPrimitives.ReadUInt16LittleEndian(body) + 1,
            BinaryPrimitives.ReadUInt16LittleEndian(body[2..]) + 1);
        int xf = pendingXf >= 0 ? pendingXf : BinaryPrimitives.ReadUInt16LittleEndian(body[4..]);
        int cursor = 6;
        string text = BiffStrings.ReadByteString(body, ref cursor, encoding, lengthBytes: 2);
        (StyleId style, _) = ResolveStyle(workbook, formats, cellFormats, xf);
        sink.OnPooledCell(cell, CellValue.Text(text), style, formula: null);
    }

    private static void ReadBoolErr(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 9)
            return;

        (CellRef cell, int xf) = ReadHeader(body, pendingXf);
        CellValue value = body[8] != 0
            ? CellValue.Error(BiffValues.DecodeError(body[7]))
            : CellValue.Boolean(body[7] != 0);
        (StyleId style, _) = ResolveStyle(workbook, formats, cellFormats, xf);
        sink.OnPooledCell(cell, value, style, formula: null);
    }

    private static void ReadFormulaValue(
        WorksheetContentSink sink, ReadOnlySpan<byte> body, ref int pendingXf,
        Workbook workbook, Dictionary<int, string> formats, List<XlsXfRecord> cellFormats)
    {
        if (body.Length < 15)
            return;

        (CellRef cell, int xf) = ReadHeader(body, pendingXf);
        BiffValues.TryReadCachedValue(body[7..15], out CellValue value);
        (StyleId style, TemporalKind temporal) = ResolveStyle(workbook, formats, cellFormats, xf);
        sink.OnPooledCell(cell, Retype(value, temporal, workbook.DateSystem), style, formula: null);
    }

    private static void Set(
        WorksheetContentSink sink, Workbook workbook, Dictionary<int, string> formats,
        List<XlsXfRecord> cellFormats, CellRef cell, int xf, double number)
    {
        if (!SheetLimits.IsValidRow(cell.Row) || !SheetLimits.IsValidColumn(cell.Column))
            return;

        (StyleId style, TemporalKind temporal) = ResolveStyle(workbook, formats, cellFormats, xf);
        sink.OnPooledCell(cell, Retype(CellValue.Number(number), temporal, workbook.DateSystem), style, formula: null);
    }

    /// <summary>Reinterprets a number as a date or a duration when its format says so.</summary>
    private static CellValue Retype(in CellValue value, TemporalKind temporal, ExcelDateSystem dateSystem)
    {
        if (temporal == TemporalKind.None || value.Kind != CellValueKind.Number)
            return value;

        double number = value.GetNumber();
        return temporal switch
        {
            TemporalKind.Date => TemporalValue.AsDate(number, dateSystem),
            TemporalKind.Duration => TemporalValue.AsDuration(number),
            _ => value,
        };
    }

    private static Encoding ResolveFallback(int codePage) => ResolveCodePage(codePage, Encoding.Latin1);

    private static Encoding ResolveCodePage(int codePage, Encoding fallback)
    {
        if (codePage is 0 or 1200)
            return fallback;

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException)
        {
            return fallback;
        }
    }
}
