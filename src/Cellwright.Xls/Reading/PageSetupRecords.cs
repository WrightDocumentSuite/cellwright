using System.Buffers.Binary;
using Cellwright.Xls.Biff;

namespace Cellwright.Xls.Reading;

/// <summary>
/// Collects the print settings of one sheet. They are scattered over a dozen records — the
/// paper in <c>Setup</c>, each margin in its own record, the fit-to-page flag two blocks
/// earlier in <c>WsBool</c> — so the pieces are gathered while the substream is read and
/// applied to the model together at its end.
/// </summary>
internal sealed class PageSetupCollector
{
    /// <summary>Margins a BIFF file means when it leaves the records out (Excel's own defaults).</summary>
    private const double DefaultSideMargin = 0.75;
    private const double DefaultTopMargin = 1.0;
    private const double DefaultHeaderMargin = 0.5;

    private double? _left;
    private double? _right;
    private double? _top;
    private double? _bottom;
    private double? _headerMargin;
    private double? _footerMargin;

    private string? _header;
    private string? _footer;
    private string? _headerEven;
    private string? _footerEven;
    private string? _headerFirst;
    private string? _footerFirst;

    private bool _centerHorizontally;
    private bool _centerVertically;
    private bool _printGridLines;
    private bool _printHeadings;
    private bool _fitToPage;
    private bool _autoPageBreaks = true;
    private bool _scaleHeaderWithDocument = true;
    private bool _alignHeaderWithMargins = true;
    private bool _differentOddEven;
    private bool _differentFirstPage;

    private SetupFields? _setup;
    private List<int>? _rowBreaks;
    private List<int>? _columnBreaks;

    /// <summary>Whether any record of the print block turned up at all.</summary>
    private bool _seenAnything;

    /// <summary>What the <c>Setup</c> record said, kept until the fit-to-page flag is known.</summary>
    private readonly record struct SetupFields(
        int PaperSize,
        int Scale,
        int FirstPage,
        int FitWidth,
        int FitHeight,
        PageOrientation Orientation,
        PageOrder Order,
        bool BlackAndWhite,
        bool Draft);

    /// <summary>Reads <c>WsBool</c>, which carries the fit-to-page and automatic-break flags.</summary>
    public void OnWsBool(ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body);
        _autoPageBreaks = (flags & 0x0001) != 0;
        _fitToPage = (flags & 0x0100) != 0;
        _seenAnything = true;
    }

    /// <summary>Reads the <c>Setup</c> record: paper, scaling, orientation and the two page counts.</summary>
    public void OnSetup(ReadOnlySpan<byte> body)
    {
        if (body.Length < 12)
            return;

        int paper = BinaryPrimitives.ReadUInt16LittleEndian(body);
        int scale = BinaryPrimitives.ReadUInt16LittleEndian(body[2..]);
        int firstPage = BinaryPrimitives.ReadInt16LittleEndian(body[4..]);
        int fitWidth = BinaryPrimitives.ReadUInt16LittleEndian(body[6..]);
        int fitHeight = BinaryPrimitives.ReadUInt16LittleEndian(body[8..]);
        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body[10..]);

        // fNoPls says the printer settings block was never filled in, and with it the paper,
        // the scale and the orientation are meaningless rather than merely default.
        bool noPrinterSettings = (flags & 0x0004) != 0;
        bool noOrientation = (flags & 0x0040) != 0;
        bool portrait = (flags & 0x0002) != 0;

        _setup = new SetupFields(
            PaperSize: noPrinterSettings ? 0 : paper,
            Scale: noPrinterSettings || scale == 0 ? 100 : scale,
            FirstPage: (flags & 0x0080) != 0 ? firstPage : 0,
            FitWidth: fitWidth,
            FitHeight: fitHeight,
            Orientation: noPrinterSettings || noOrientation
                ? PageOrientation.Default
                : portrait ? PageOrientation.Portrait : PageOrientation.Landscape,
            Order: (flags & 0x0001) != 0 ? PageOrder.OverThenDown : PageOrder.DownThenOver,
            BlackAndWhite: (flags & 0x0008) != 0,
            Draft: (flags & 0x0010) != 0);

        // The header and footer margins ride along in this record rather than in their own.
        if (body.Length >= 32)
        {
            _headerMargin = BinaryPrimitives.ReadDoubleLittleEndian(body[16..]);
            _footerMargin = BinaryPrimitives.ReadDoubleLittleEndian(body[24..]);
        }

        _seenAnything = true;
    }

    /// <summary>Reads one of the four margin records, each an eight-byte number of inches.</summary>
    public void OnMargin(ushort recordType, ReadOnlySpan<byte> body)
    {
        if (body.Length < 8)
            return;

        double inches = BinaryPrimitives.ReadDoubleLittleEndian(body);
        switch (recordType)
        {
            case BiffRecordType.LeftMargin:
                _left = inches;
                break;
            case BiffRecordType.RightMargin:
                _right = inches;
                break;
            case BiffRecordType.TopMargin:
                _top = inches;
                break;
            case BiffRecordType.BottomMargin:
                _bottom = inches;
                break;
            default:
                return;
        }

        _seenAnything = true;
    }

    /// <summary>Reads the <c>Header</c> or <c>Footer</c> record; an empty record means none.</summary>
    public void OnHeaderFooterText(ushort recordType, ReadOnlySpan<byte> body)
    {
        _seenAnything = true;
        if (body.Length < 3)
            return;

        int cursor = 0;
        string text = BiffStrings.ReadUnicodeString(body, ref cursor);
        if (text.Length == 0)
            return;

        if (recordType == BiffRecordType.Header)
            _header = text;
        else
            _footer = text;
    }

    /// <summary>
    /// Reads the <c>HeaderFooter</c> record, which adds the even and first-page variants that
    /// the two older records have no room for.
    /// </summary>
    public void OnHeaderFooterExtension(ReadOnlySpan<byte> body)
    {
        // 12 bytes of future-record header, then the view guid, the flags and four lengths.
        const int FlagsOffset = 28;
        if (body.Length < FlagsOffset + 10)
            return;

        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(body[FlagsOffset..]);
        _differentOddEven = (flags & 0x0001) != 0;
        _differentFirstPage = (flags & 0x0002) != 0;
        _scaleHeaderWithDocument = (flags & 0x0004) != 0;
        _alignHeaderWithMargins = (flags & 0x0008) != 0;

        // Each length is restated by the string that follows, so a length of zero means the
        // string is simply not there.
        int cursor = FlagsOffset + 2;
        bool hasEvenHeader = BinaryPrimitives.ReadUInt16LittleEndian(body[cursor..]) != 0;
        bool hasEvenFooter = BinaryPrimitives.ReadUInt16LittleEndian(body[(cursor + 2)..]) != 0;
        bool hasFirstHeader = BinaryPrimitives.ReadUInt16LittleEndian(body[(cursor + 4)..]) != 0;
        bool hasFirstFooter = BinaryPrimitives.ReadUInt16LittleEndian(body[(cursor + 6)..]) != 0;
        cursor += 8;

        _headerEven = hasEvenHeader ? BiffStrings.ReadUnicodeString(body, ref cursor) : null;
        _footerEven = hasEvenFooter ? BiffStrings.ReadUnicodeString(body, ref cursor) : null;
        _headerFirst = hasFirstHeader ? BiffStrings.ReadUnicodeString(body, ref cursor) : null;
        _footerFirst = hasFirstFooter ? BiffStrings.ReadUnicodeString(body, ref cursor) : null;
        _seenAnything = true;
    }

    /// <summary>Reads one of the four single-flag records: centering, gridlines, headings.</summary>
    public void OnPrintFlag(ushort recordType, ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        bool set = BinaryPrimitives.ReadUInt16LittleEndian(body) != 0;
        switch (recordType)
        {
            case BiffRecordType.HCenter:
                _centerHorizontally = set;
                break;
            case BiffRecordType.VCenter:
                _centerVertically = set;
                break;
            case BiffRecordType.PrintGrid:
                _printGridLines = set;
                break;
            case BiffRecordType.PrintRowCol:
                _printHeadings = set;
                break;
            default:
                return;
        }

        _seenAnything = true;
    }

    /// <summary>
    /// Reads a page-break list. Each entry names the row below or the column right of the
    /// break, together with the span it covers, which the model does not keep.
    /// </summary>
    public void OnPageBreaks(ushort recordType, ReadOnlySpan<byte> body)
    {
        if (body.Length < 2)
            return;

        int count = BinaryPrimitives.ReadUInt16LittleEndian(body);
        var positions = new List<int>(Math.Min(count, 1026));
        for (int i = 0; i < count && 2 + (i * 6) + 6 <= body.Length; i++)
        {
            int position = BinaryPrimitives.ReadUInt16LittleEndian(body[(2 + (i * 6))..]) + 1;
            if (position > 1)
                positions.Add(position);
        }

        if (positions.Count == 0)
            return;

        positions.Sort();
        if (recordType == BiffRecordType.HorizontalPageBreaks)
            _rowBreaks = positions;
        else
            _columnBreaks = positions;
        _seenAnything = true;
    }

    /// <summary>Applies everything collected to the sheet.</summary>
    /// <remarks>
    /// The settings become the model's own rather than the file's: a workbook read from BIFF
    /// has no markup to replay on save, so the modern writer only emits what the model owns.
    /// </remarks>
    public void Finish(Worksheet sheet)
    {
        if (!_seenAnything)
            return;

        PageMargins? margins = BuildMargins();
        HeaderFooterText? header = Lines(_header, _headerEven, _headerFirst);
        HeaderFooterText? footer = Lines(_footer, _footerEven, _footerFirst);

        Apply(sheet.PageSetup, target =>
        {
            target.AutoPageBreaks = _autoPageBreaks;
            target.CenterHorizontally = _centerHorizontally;
            target.CenterVertically = _centerVertically;
            target.PrintGridLines = _printGridLines;
            target.PrintHeadings = _printHeadings;
            target.ScaleHeaderWithDocument = _scaleHeaderWithDocument;
            target.AlignHeaderWithMargins = _alignHeaderWithMargins;
            target.DifferentOddEven = _differentOddEven;
            target.DifferentFirstPage = _differentFirstPage;

            if (margins is not null)
                target.Margins = margins;
            if (header is not null)
                target.Header = header;
            if (footer is not null)
                target.Footer = footer;

            if (_fitToPage)
                target.FitToPageRequested = true;

            if (_setup is not { } setup)
                return;

            target.PaperSize = setup.PaperSize;
            target.Scale = setup.Scale;
            target.FirstPageNumber = setup.FirstPage;
            target.Orientation = setup.Orientation;
            target.Order = setup.Order;
            target.BlackAndWhite = setup.BlackAndWhite;
            target.Draft = setup.Draft;

            // The two page counts sit in Setup whether or not fitting is on, and mean nothing
            // without the flag that WsBool carries.
            if (_fitToPage)
            {
                target.FitToWidth = setup.FitWidth == 0 ? null : setup.FitWidth;
                target.FitToHeight = setup.FitHeight == 0 ? null : setup.FitHeight;
            }
        });

        if (_rowBreaks is not null || _columnBreaks is not null)
        {
            sheet.PageSetup.SetBreaksFromFile(_rowBreaks, _columnBreaks);
            sheet.PageSetup.MarkModified();
        }
    }

    /// <summary>Fills in the settings and hands ownership of them to the model.</summary>
    private static void Apply(PageSetup setup, Action<PageSetup> fill)
    {
        fill(setup);
        setup.MarkModified();
    }

    /// <summary>
    /// Builds the margins from whichever records turned up. A file that named none of them is
    /// left alone: the model would rather say nothing than claim Excel's defaults were chosen.
    /// </summary>
    private PageMargins? BuildMargins()
    {
        if (_left is null && _right is null && _top is null && _bottom is null)
            return null;

        return new PageMargins(
            _left ?? DefaultSideMargin,
            _right ?? DefaultSideMargin,
            _top ?? DefaultTopMargin,
            _bottom ?? DefaultTopMargin,
            _headerMargin ?? DefaultHeaderMargin,
            _footerMargin ?? DefaultHeaderMargin);
    }

    private static HeaderFooterText? Lines(string? standard, string? even, string? first)
    {
        var text = new HeaderFooterText { Default = standard, Even = even, First = first };
        return text.IsEmpty ? null : text;
    }
}
