namespace Cellwright.Xls.Xlsb;

/// <summary>
/// The record types this reader acts on, by the numbers of the record enumeration
/// (MS-XLSB 2.3). The format defines well over a thousand; these are the ones that carry
/// something the model holds.
/// </summary>
/// <remarks>
/// The single-byte numbers below 128 are not an accident of ordering: the format reserves them
/// for the records a sheet is mostly made of, so a cell costs one byte to name rather than two.
/// </remarks>
internal static class Biff12RecordType
{
    // Cells and rows, the records a worksheet is almost entirely made of.
    public const int RowHdr = 0;
    public const int CellBlank = 1;
    public const int CellRk = 2;
    public const int CellError = 3;
    public const int CellBool = 4;
    public const int CellReal = 5;
    public const int CellSt = 6;
    public const int CellIsst = 7;
    public const int FmlaString = 8;
    public const int FmlaNum = 9;
    public const int FmlaBool = 10;
    public const int FmlaError = 11;
    public const int CellRString = 62;

    // Shared strings.
    public const int SstItem = 19;
    public const int BeginSst = 159;
    public const int EndSst = 160;

    // Styles.
    public const int Font = 43;
    public const int Fmt = 44;
    public const int Fill = 45;
    public const int Border = 46;
    public const int Xf = 47;
    public const int Style = 48;
    public const int BeginFills = 603;
    public const int EndFills = 604;
    public const int BeginFonts = 611;
    public const int EndFonts = 612;
    public const int BeginBorders = 613;
    public const int EndBorders = 614;
    public const int BeginFmts = 615;
    public const int EndFmts = 616;
    public const int BeginCellXFs = 617;
    public const int EndCellXFs = 618;
    public const int BeginCellStyleXFs = 626;
    public const int EndCellStyleXFs = 627;
    public const int BeginStyleSheet = 278;
    public const int EndStyleSheet = 279;

    // The workbook part.
    public const int Name = 39;
    public const int BeginBook = 131;
    public const int EndBook = 132;
    public const int BeginBundleShs = 143;
    public const int EndBundleShs = 144;
    public const int WbProp = 153;
    public const int BundleSh = 156;
    public const int BeginExternals = 353;
    public const int EndExternals = 354;
    public const int SupBookSrc = 355;
    public const int SupSelf = 357;
    public const int SupSame = 358;
    public const int SupTabs = 359;
    public const int ExternSheet = 362;
    public const int SupAddin = 667;

    // Worksheet structure.
    public const int BeginSheet = 129;
    public const int EndSheet = 130;
    public const int BeginWsView = 137;
    public const int BeginSheetData = 145;
    public const int EndSheetData = 146;
    public const int WsDim = 148;
    public const int MergeCell = 176;
    public const int BeginMergeCells = 177;
    public const int EndMergeCells = 178;
    public const int ColInfo = 60;
    public const int BeginColInfos = 390;
    public const int EndColInfos = 391;
    public const int HLink = 494;
    public const int DVal = 64;
    public const int BeginDVals = 573;
    public const int EndDVals = 574;

    // Formulas that several cells share.
    public const int ArrFmla = 426;
    public const int ShrFmla = 427;

    // Tables, whose names structured references are written against.
    public const int BeginList = 343;
    public const int EndList = 344;
    public const int BeginListCols = 345;
    public const int EndListCols = 346;
    public const int BeginListCol = 347;
    public const int EndListCol = 348;

    // Notes, which live in a part of their own beside the sheet.
    public const int BeginComments = 628;
    public const int EndComments = 629;
    public const int BeginCommentAuthors = 630;
    public const int EndCommentAuthors = 631;
    public const int CommentAuthor = 632;
    public const int BeginCommentList = 633;
    public const int EndCommentList = 634;
    public const int BeginComment = 635;
    public const int EndComment = 636;
    public const int CommentText = 637;

    // Conditional formatting.
    public const int BeginConditionalFormatting = 461;
    public const int EndConditionalFormatting = 462;
    public const int BeginCFRule = 463;
    public const int EndCFRule = 464;
    public const int BeginIconSet = 465;
    public const int EndIconSet = 466;
    public const int BeginDatabar = 467;
    public const int EndDatabar = 468;
    public const int BeginColorScale = 469;
    public const int EndColorScale = 470;
    public const int CFVO = 471;
    public const int Color = 564;
}
