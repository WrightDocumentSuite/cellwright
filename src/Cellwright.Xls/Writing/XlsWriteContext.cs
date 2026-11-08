namespace Cellwright.Xls.Writing;

/// <summary>
/// The tables the writer builds up front and every part of the write then indexes into: the
/// shared strings, the style table, the sheet list and the defined names.
/// </summary>
/// <remarks>
/// It exists because BIFF is written back to front — a formula names a sheet by its position in
/// a table that appears earlier in the file, so every such table has to be complete before the
/// first cell is written.
/// </remarks>
internal sealed class XlsWriteContext(Workbook workbook, Action<WorkbookWarning> warn)
{
    private readonly Dictionary<string, int> _sharedStrings = new(StringComparer.Ordinal);
    private readonly Dictionary<(string Text, string Markup, int DefaultFont), int> _sharedRichStrings = [];
    private readonly List<string> _sharedStringList = [];
    private readonly List<(ushort CharPos, ushort FontIndex)[]?> _sharedStringRuns = [];
    private readonly Dictionary<string, int> _externSheets = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<int> _externSheetTargets = [];
    private readonly Dictionary<string, int> _nameIndexes = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The workbook being written.</summary>
    public Workbook Workbook => workbook;

    /// <summary>
    /// Which generation of the format is being written. It decides how a formula's references are
    /// laid out and how large a grid they may name — the two things that differ between the
    /// record streams rather than merely being framed differently.
    /// </summary>
    public Reading.BiffVersion Version { get; init; } = Reading.BiffVersion.Biff8;

    /// <summary>
    /// The style table, for the parts of a sheet that need a palette slot of their own — a
    /// conditional formatting rule paints in colours no cell style names.
    /// </summary>
    public required StyleTable StyleTable { get; init; }

    /// <summary>Reports something the format cannot carry.</summary>
    public void Warn(WorkbookWarning warning) => warn(warning);

    /// <summary>Total number of times a shared string was referenced, which the SST records.</summary>
    public int TotalStringReferences { get; private set; }

    /// <summary>The unique shared strings, in the order the SST will write them.</summary>
    public IReadOnlyList<string> SharedStrings => _sharedStringList;

    /// <summary>Character-formatting runs parallel to <see cref="SharedStrings"/>.</summary>
    public IReadOnlyList<(ushort CharPos, ushort FontIndex)[]?> SharedStringRuns => _sharedStringRuns;

    /// <summary>The sheets each ExternSheet entry points at, in order.</summary>
    public IReadOnlyList<int> ExternSheetTargets => _externSheetTargets;

    /// <summary>Style handles mapped to XF indexes in the file's own table.</summary>
    public required IReadOnlyDictionary<int, int> XfByStyle { get; init; }

    /// <summary>Whether a style index carries a date or duration format, deciding how to write a value.</summary>
    public required IReadOnlyList<Styles.TemporalKind> TemporalByXf { get; init; }

    /// <summary>Interns a string into the shared table and returns its index.</summary>
    public int InternString(string text)
    {
        TotalStringReferences++;
        if (_sharedStrings.TryGetValue(text, out int index))
            return index;

        index = _sharedStringList.Count;
        _sharedStrings[text] = index;
        _sharedStringList.Add(text);
        _sharedStringRuns.Add(null);
        return index;
    }

    /// <summary>Interns a rich-text string and its formatting runs into the shared table.</summary>
    public int InternRichString(RichText richText, string markup, int defaultFontIndex)
    {
        TotalStringReferences++;
        string text = richText.Text;
        var key = (text, markup, defaultFontIndex);
        if (_sharedRichStrings.TryGetValue(key, out int index))
            return index;

        var runs = new List<(ushort CharPos, ushort FontIndex)>(richText.Count);
        int charPos = 0;
        foreach (RichTextRun run in richText)
        {
            if (run.Text.Length == 0)
                continue;

            int fontIndex = run.Font is { } font
                ? StyleTable.FontIndex(font)
                : defaultFontIndex;
            runs.Add(((ushort)charPos, (ushort)fontIndex));
            charPos += run.Text.Length;
        }

        index = _sharedStringList.Count;
        _sharedRichStrings[key] = index;
        _sharedStringList.Add(text);
        _sharedStringRuns.Add(runs.Count > 0 ? [.. runs] : null);
        return index;
    }

    /// <summary>Returns the XF index a style handle maps to, or the default format.</summary>
    public int XfFor(int styleIndex) => XfByStyle.TryGetValue(styleIndex, out int xf) ? xf : DefaultCellXf;

    /// <summary>Index of the first cell XF, which follows the fifteen style XFs Excel expects.</summary>
    public const int DefaultCellXf = 15;

    /// <summary>Whether an XF index formats its number as a date or a duration.</summary>
    public Styles.TemporalKind TemporalOf(int xfIndex) =>
        (uint)xfIndex < (uint)TemporalByXf.Count ? TemporalByXf[xfIndex] : Styles.TemporalKind.None;

    /// <summary>
    /// Returns the ExternSheet index of a sheet by name, adding one on first use. A formula
    /// naming another sheet points at this table rather than at the sheet directly.
    /// </summary>
    public int? ExternSheetFor(string sheetName)
    {
        if (_externSheets.TryGetValue(sheetName, out int index))
            return index;

        int sheet = IndexOfSheet(sheetName);
        if (sheet < 0)
            return null;

        index = _externSheetTargets.Count;
        _externSheets[sheetName] = index;
        _externSheetTargets.Add(sheet);
        return index;
    }

    /// <summary>Returns the zero-based index of a defined name, or null when there is none.</summary>
    public int? NameIndexFor(string name) => _nameIndexes.TryGetValue(name, out int index) ? index : null;

    /// <summary>Records where a defined name sits in the Lbl table.</summary>
    public void RegisterName(string name, int index) => _nameIndexes[name] = index;

    /// <summary>
    /// The drawings the sheets created, in order. Notes live in a drawing, and the workbook-wide
    /// record that has to account for every shape is written after the sheets that made them.
    /// </summary>
    public IReadOnlyList<(int DrawingId, int ShapeCount)> Drawings => _drawings;

    private readonly List<(int DrawingId, int ShapeCount)> _drawings = [];

    /// <summary>Claims the next drawing id, which decides the range its shape ids come from.</summary>
    public int NextDrawingId() => _drawings.Count + 1;

    /// <summary>Records how many shapes a drawing ended up with.</summary>
    public void RegisterDrawing(int drawingId, int shapeCount) => _drawings.Add((drawingId, shapeCount));

    /// <summary>
    /// The images of the workbook's pictures, in store order. The store is workbook-wide while
    /// the shapes that use it are per-sheet, which is why it lives here: a sheet registers its
    /// images while being written, and the globals emit the store afterwards.
    /// </summary>
    public IReadOnlyList<(ReadOnlyMemory<byte> Content, PictureFormat Format, int References)> Blips => _blips;

    private readonly List<(ReadOnlyMemory<byte> Content, PictureFormat Format, int References)> _blips = [];
    private readonly Dictionary<int, int> _blipByHash = [];

    /// <summary>
    /// Places an image in the store, reusing the entry when the same bytes are already there,
    /// and returns its 1-based index — the value the shape's <c>pib</c> property carries.
    /// </summary>
    public int RegisterBlip(ReadOnlyMemory<byte> content, PictureFormat format)
    {
        // Content-hash dedup, the same idea the modern writer applies to image parts: ten
        // shapes of one logo cost one store entry with a reference count of ten.
        var hash = new HashCode();
        hash.AddBytes(content.Span);
        int key = hash.ToHashCode();

        if (_blipByHash.TryGetValue(key, out int index) &&
            _blips[index - 1].Content.Span.SequenceEqual(content.Span))
        {
            (ReadOnlyMemory<byte> stored, PictureFormat storedFormat, int references) = _blips[index - 1];
            _blips[index - 1] = (stored, storedFormat, references + 1);
            return index;
        }

        _blips.Add((content, format, 1));
        _blipByHash[key] = _blips.Count;
        return _blips.Count;
    }

    private int IndexOfSheet(string name)
    {
        for (int i = 0; i < workbook.Sheets.Count; i++)
        {
            if (string.Equals(workbook.Sheets[i].Name, name, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }
}
