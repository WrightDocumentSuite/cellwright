using System.Text;
using Cellwright.Styles;

namespace Cellwright.Xls.Reading;

/// <summary>How a sheet appears in the workbook's tab strip, as BoundSheet8 states it.</summary>
internal enum XlsSheetKind
{
    /// <summary>An ordinary grid.</summary>
    Worksheet,

    /// <summary>A chart on a tab of its own.</summary>
    Chart,

    /// <summary>A macro sheet of either flavour, or a dialog sheet.</summary>
    Other,
}

/// <summary>One entry of the workbook's sheet list.</summary>
/// <param name="Name">Tab name.</param>
/// <param name="Offset">Where the sheet's substream starts in the workbook stream.</param>
/// <param name="Visibility">How visible the tab is.</param>
/// <param name="Kind">What kind of sheet it is.</param>
internal readonly record struct XlsSheetEntry(string Name, int Offset, SheetVisibility Visibility, XlsSheetKind Kind);

/// <summary>A SupBook record: one place this workbook's 3-D references can reach into.</summary>
/// <remarks>
/// Three things wear the same record. Most files carry one standing for the workbook itself,
/// whose sheet indexes are local. A second kind names another workbook on disk. A third stands
/// for the add-ins, and holds no sheets at all — only the names of the functions the workbook
/// borrows from them, which is how EOMONTH and its relatives reach a formula.
/// </remarks>
internal sealed class XlsSupBook
{
    /// <summary>
    /// Display path of the other workbook, or <see langword="null"/> when the record stands for
    /// this workbook or for the add-ins.
    /// </summary>
    public string? Path { get; init; }

    /// <summary>Tab names of the other workbook, which its sheet indexes count.</summary>
    public IReadOnlyList<string> SheetNames { get; init; } = [];

    /// <summary>Whether the record stands for the add-ins rather than a workbook.</summary>
    public bool IsAddin { get; init; }

    /// <summary>
    /// Names borrowed from this book, in the order of the ExternName records that follow it.
    /// A PtgNameX indexes into this list, one-based.
    /// </summary>
    public List<string> Names { get; } = [];

    /// <summary>Whether the record points somewhere other than the workbook holding it.</summary>
    public bool IsExternal => Path is not null;
}

/// <summary>A defined name as the Lbl record states it.</summary>
/// <param name="Name">Name as used in formulas.</param>
/// <param name="Rgce">The encoded expression it stands for.</param>
/// <param name="SheetIndex">One-based sheet the name is local to, or 0 for the whole workbook.</param>
/// <param name="IsHidden">Whether the name manager hides it.</param>
/// <param name="IsBuiltIn">Whether the name is one of Excel's own, whose name is a single byte code.</param>
internal readonly record struct XlsDefinedName(
    string Name, byte[] Rgce, int SheetIndex, bool IsHidden, bool IsBuiltIn);

/// <summary>
/// Everything the global substream of a workbook says: the sheet list, the shared strings, the
/// style tables and the defined names. Read once, then handed to every sheet parser.
/// </summary>
internal sealed class XlsGlobals
{
    /// <summary>BIFF version the file declares.</summary>
    public BiffVersion Version { get; set; } = BiffVersion.Biff8;

    /// <summary>Date system the workbook counts from.</summary>
    public ExcelDateSystem DateSystem { get; set; }

    /// <summary>Encoding of the byte strings of BIFF2-5.</summary>
    public Encoding Encoding { get; set; } = Encoding.Latin1;

    /// <summary>The sheets, in tab order.</summary>
    public List<XlsSheetEntry> Sheets { get; } = [];

    /// <summary>The shared-string table, indexed by LabelSst.</summary>
    public string[] SharedStrings { get; set; } = [];

    /// <summary>Rich-text runs parallel to <see cref="SharedStrings"/>, when an entry carries character formatting.</summary>
    public RichText?[] SharedRichStrings { get; set; } = [];

    /// <summary>Resolved workbook fonts in table order.</summary>
    public List<FontStyle> Fonts { get; } = [];

    /// <summary>Cell style handles by XF index, mapped into the workbook's pool.</summary>
    public List<StyleId> CellFormats { get; } = [];

    /// <summary>Whether each XF index formats a date or a duration, which decides how a number reads.</summary>
    public List<TemporalKind> TemporalKinds { get; } = [];

    /// <summary>Defined names, in the order the file lists them.</summary>
    public List<XlsDefinedName> Names { get; } = [];

    /// <summary>
    /// Entries of the ExternSheet record, which 3-D references point into. Entry <c>i</c> names
    /// the workbook it belongs to and the first and last sheet of reference <c>i</c>; a sheet
    /// index of -1 means the tab was deleted.
    /// </summary>
    /// <remarks>
    /// The workbook index matters: for an entry of another workbook the sheet indexes count
    /// tabs of <em>that</em> book, so reading them against this one silently names a different
    /// sheet that happens to sit at the same position.
    /// </remarks>
    public List<(int SupBook, int First, int Last)> ExternSheets { get; } = [];

    /// <summary>External workbooks, in SupBook order, for rendering 3-D references.</summary>
    public List<XlsSupBook> SupBooks { get; } = [];

    /// <summary>
    /// Every image in the file, in the order the drawing group stored them. A sheet's shapes
    /// name one by its position here rather than carrying the bytes themselves, which is why
    /// this is workbook-wide: the same picture on three sheets is stored once.
    /// </summary>
    public List<EscherPictures.Blip> Blips { get; } = [];

    /// <summary>
    /// The workbook's color table, if it overrode the standard one. The legacy format has no
    /// notion of a theme: every color is a slot in a table of 56 plus a few fixed meanings.
    /// </summary>
    public List<uint> Palette { get; } = [];

    /// <summary>Resolves a palette index to a color.</summary>
    public ExcelColor ColorOf(int index)
    {
        // 0x7FFF is "automatic"; 64 and 65 are the system window colors, which have no value.
        if (index is 0x7FFF or 64 or 65 or 0)
            return ExcelColor.Auto;

        int slot = index - 8;
        return (uint)slot < (uint)Palette.Count ? ExcelColor.FromArgb(Palette[slot]) : ExcelColor.FromIndexed(index);
    }

    /// <summary>Resolves an XF index to a style handle, falling back to the default.</summary>
    public StyleId StyleOf(int xfIndex) =>
        (uint)xfIndex < (uint)CellFormats.Count ? CellFormats[xfIndex] : StyleId.Default;

    /// <summary>Whether an XF index formats its number as a date or a duration.</summary>
    public TemporalKind TemporalOf(int xfIndex) =>
        (uint)xfIndex < (uint)TemporalKinds.Count ? TemporalKinds[xfIndex] : TemporalKind.None;

    /// <summary>Looks up a shared string, tolerating an index the file got wrong.</summary>
    public string StringAt(int index) =>
        (uint)index < (uint)SharedStrings.Length ? SharedStrings[index] : "";

    /// <summary>Looks up the rich-text runs of a shared string, or <see langword="null"/> when plain.</summary>
    public RichText? RichStringAt(int index) =>
        (uint)index < (uint)SharedRichStrings.Length ? SharedRichStrings[index] : null;

    /// <summary>Resolves a font index from the workbook's font table.</summary>
    public FontStyle FontOf(int fontIndex)
    {
        int slot = Version != BiffVersion.Biff12 && fontIndex > 4 ? fontIndex - 1 : fontIndex;
        return (uint)slot < (uint)Fonts.Count ? Fonts[slot] : FontStyle.Default;
    }

    /// <summary>Builds a <see cref="RichText"/> value from character-offset formatting runs.</summary>
    public RichText? BuildRichText(string text, ReadOnlySpan<(int CharPos, int FontIndex)> runs)
    {
        if (runs.IsEmpty || text.Length == 0)
            return null;

        var result = new List<RichTextRun>(runs.Length + 1);
        int firstStart = Math.Clamp(runs[0].CharPos, 0, text.Length);
        if (firstStart > 0)
            result.Add(new RichTextRun(text[..firstStart]));

        for (int i = 0; i < runs.Length; i++)
        {
            int start = Math.Clamp(runs[i].CharPos, 0, text.Length);
            int end = i + 1 < runs.Length
                ? Math.Clamp(runs[i + 1].CharPos, start, text.Length)
                : text.Length;
            if (end <= start)
                continue;

            result.Add(new RichTextRun(text[start..end], FontOf(runs[i].FontIndex)));
        }

        if (result.Count == 0)
            return null;

        var rich = new RichText(result);
        return rich.IsFormatted ? rich : null;
    }
}

/// <summary>The BIFF versions this library can read.</summary>
internal enum BiffVersion
{
    /// <summary>Excel 2.x, a bare record stream with no container.</summary>
    Biff2 = 2,

    /// <summary>Excel 3.0.</summary>
    Biff3 = 3,

    /// <summary>Excel 4.0.</summary>
    Biff4 = 4,

    /// <summary>Excel 5.0 and 95, the first to use a compound file.</summary>
    Biff5 = 5,

    /// <summary>Excel 97 through 2003, the last and richest of the binary formats.</summary>
    Biff8 = 8,

    /// <summary>
    /// The records of a binary .xlsb workbook. Not a successor to BIFF8 in the usual sense —
    /// the container is a zip package rather than a compound file, and the records are framed
    /// differently — but the formula tokens are nearly the same, which is what this value is
    /// for: it tells the decoder that a reference carries a four-byte row.
    /// </summary>
    Biff12 = 12,
}
