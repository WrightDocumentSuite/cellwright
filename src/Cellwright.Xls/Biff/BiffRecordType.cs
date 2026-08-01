namespace Cellwright.Xls.Biff;

/// <summary>
/// The record types this library reacts to. A BIFF stream holds hundreds more; anything absent
/// here is skipped by size, which is what makes a forward-only reader tolerant of files written
/// by producers nobody has heard of.
/// </summary>
internal static class BiffRecordType
{
    // Stream and substream structure.
    public const ushort Formula = 0x0006;
    public const ushort Eof = 0x000A;
    public const ushort CalcCount = 0x000C;
    public const ushort Date1904 = 0x0022;
    public const ushort Continue = 0x003C;
    public const ushort Bof = 0x0809;
    public const ushort Interfacehdr = 0x00E1;
    public const ushort FilePass = 0x002F;
    public const ushort CodePage = 0x0042;

    // Globals.
    public const ushort BoundSheet8 = 0x0085;
    public const ushort Sst = 0x00FC;
    public const ushort ExtSst = 0x00FF;
    public const ushort Font = 0x0031;
    public const ushort Format = 0x041E;
    public const ushort Xf = 0x00E0;
    public const ushort Style = 0x0293;
    public const ushort Palette = 0x0092;
    public const ushort Lbl = 0x0018;
    public const ushort SupBook = 0x01AE;
    public const ushort ExternSheet = 0x0017;
    public const ushort ExternName = 0x0023;
    public const ushort Window1 = 0x003D;

    // Cell table.
    public const ushort Row = 0x0208;
    public const ushort Blank = 0x0201;
    public const ushort MulBlank = 0x00BE;
    public const ushort Rk = 0x027E;
    public const ushort MulRk = 0x00BD;
    public const ushort Number = 0x0203;
    public const ushort BoolErr = 0x0205;
    public const ushort LabelSst = 0x00FD;
    public const ushort Label = 0x0204;
    public const ushort RString = 0x00D6;
    public const ushort String = 0x0207;
    public const ushort ShrFmla = 0x04BC;
    public const ushort Array = 0x0221;
    public const ushort Table = 0x0236;
    public const ushort Dimensions = 0x0200;
    public const ushort Index = 0x020B;
    public const ushort DbCell = 0x00D7;

    // Sheet layout and features.
    public const ushort ColInfo = 0x007D;
    public const ushort DefaultRowHeight = 0x0225;
    public const ushort DefColWidth = 0x0055;
    public const ushort Window2 = 0x023E;
    public const ushort Pane = 0x0041;
    public const ushort Scl = 0x00A0;
    public const ushort MergeCells = 0x00E5;
    public const ushort HLink = 0x01B8;
    public const ushort HLinkTooltip = 0x0800;
    public const ushort Dv = 0x01BE;
    public const ushort DVal = 0x01B2;
    public const ushort SheetExt = 0x0862;
    public const ushort Protect = 0x0012;
    public const ushort Password = 0x0013;
    public const ushort AutoFilterInfo = 0x009D;
    public const ushort Note = 0x001C;
    public const ushort Obj = 0x005D;
    public const ushort TxO = 0x01B6;
    public const ushort Guts = 0x0080;
    public const ushort Selection = 0x001D;

    // Printing: the settings are spread over a dozen single-purpose records.
    public const ushort Setup = 0x00A1;
    public const ushort LeftMargin = 0x0026;
    public const ushort RightMargin = 0x0027;
    public const ushort TopMargin = 0x0028;
    public const ushort BottomMargin = 0x0029;
    public const ushort Header = 0x0014;
    public const ushort Footer = 0x0015;
    public const ushort HeaderFooter = 0x089C;
    public const ushort HCenter = 0x0083;
    public const ushort VCenter = 0x0084;
    public const ushort PrintRowCol = 0x002A;
    public const ushort PrintGrid = 0x002B;
    public const ushort GridSet = 0x0082;
    public const ushort WsBool = 0x0081;
    public const ushort HorizontalPageBreaks = 0x001B;
    public const ushort VerticalPageBreaks = 0x001A;

    // Conditional formatting.
    public const ushort CondFmt = 0x01B0;
    public const ushort Cf = 0x01B1;
    public const ushort CondFmt12 = 0x0879;
    public const ushort Cf12 = 0x087A;

    /// <summary>Extends a classic CondFmt group with 2007-only settings; not a rule of its own.</summary>
    public const ushort CfEx = 0x087B;

    // Drawing layer, which is how notes are anchored to their cells.
    public const ushort MsoDrawingGroup = 0x00EB;
    public const ushort MsoDrawing = 0x00EC;
    public const ushort MsoDrawingSelection = 0x00ED;

    // Calculation settings, written because the sheet globals block expects them.
    public const ushort CalcMode = 0x000D;
    public const ushort CalcRefMode = 0x000F;
    public const ushort CalcDelta = 0x0010;
    public const ushort CalcIter = 0x0011;
    public const ushort CalcSaveRecalc = 0x005F;

    // Legacy shapes of records that changed between versions.
    public const ushort OldFormat = 0x001E;   // FORMAT in BIFF2-5.
    public const ushort OldFont = 0x0231;     // FONT in BIFF3-4.
    public const ushort Biff2Font = 0x0031;   // Same id as BIFF8 FONT but a different body.
    public const ushort Biff2Row = 0x0008;
    public const ushort Biff2Blank = 0x0001;
    public const ushort Biff2Integer = 0x0002;
    public const ushort Biff2Number = 0x0003;
    public const ushort Biff2Label = 0x0004;
    public const ushort Biff2BoolErr = 0x0005;
    public const ushort Biff2Formula = 0x0006;
    public const ushort Biff2Xf = 0x0043;
    public const ushort Biff3Xf = 0x0243;
    public const ushort Biff4Xf = 0x0443;
    public const ushort Ixfe = 0x0044;
}
