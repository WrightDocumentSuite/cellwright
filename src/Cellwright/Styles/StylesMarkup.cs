namespace Cellwright.Styles;

/// <summary>
/// Sections of <c>styles.xml</c> the style model does not represent. They are carried through
/// verbatim because other parts index into them: a table column names a differential format
/// by position, and losing <c>dxfs</c> leaves the table pointing at nothing.
/// </summary>
internal sealed class StylesMarkup
{
    /// <summary>Extra attributes of the root element, e.g. <c>mc:Ignorable</c> and namespace declarations.</summary>
    public List<KeyValuePair<string, string>>? RootAttributes { get; set; }

    /// <summary>
    /// Content of the <c>dxfs</c> element: the differential formats tables and conditional
    /// formatting index into. Stored without its wrapper so that formats the model adds can be
    /// appended behind them, keeping every index the file already used.
    /// </summary>
    public string? DifferentialFormats { get; set; }

    /// <summary>How many formats <see cref="DifferentialFormats"/> holds, i.e. the first free index.</summary>
    public int DifferentialFormatCount { get; set; }

    /// <summary>
    /// The master formats of <c>cellStyleXfs</c>, parsed rather than kept raw: their component
    /// ids point into tables the writer rebuilds, so replaying the bytes would leave them
    /// indexing someone else's font. Order matters — <c>cellStyles</c> names them by position.
    /// </summary>
    public IReadOnlyList<CellStyle>? StyleFormats { get; set; }

    /// <summary>The <c>cellStyles</c> element naming the entries of <see cref="StyleFormats"/>.</summary>
    public string? CellStyleEntries { get; set; }

    /// <summary>The <c>tableStyles</c> element.</summary>
    public string? TableStyles { get; set; }

    /// <summary>The <c>colors</c> element: the indexed palette and MRU colors.</summary>
    public string? Colors { get; set; }

    /// <summary>The trailing <c>extLst</c> element.</summary>
    public string? Extensions { get; set; }

    /// <summary>Whether anything at all was captured.</summary>
    public bool IsEmpty =>
        RootAttributes is null && DifferentialFormats is null &&
        StyleFormats is null && CellStyleEntries is null &&
        TableStyles is null && Colors is null && Extensions is null;
}
