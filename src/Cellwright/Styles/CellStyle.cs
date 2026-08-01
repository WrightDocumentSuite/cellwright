namespace Cellwright;

/// <summary>
/// A complete immutable cell style. Styles are plain records: build them with
/// <c>with</c> expressions (or the fluent extensions) and the library interns
/// equal styles into a single format entry when writing.
/// </summary>
public sealed record CellStyle
{
    /// <summary>The workbook default style.</summary>
    public static CellStyle Default { get; } = new();

    /// <summary>Font settings.</summary>
    public FontStyle Font { get; init; } = FontStyle.Default;

    /// <summary>Fill settings.</summary>
    public FillStyle Fill { get; init; } = FillStyle.None;

    /// <summary>Border settings.</summary>
    public BorderStyle Border { get; init; } = BorderStyle.None;

    /// <summary>Alignment settings.</summary>
    public Alignment Alignment { get; init; } = Alignment.Default;

    /// <summary>What happens to the cell once its sheet is protected.</summary>
    public CellProtection Protection { get; init; } = CellProtection.Default;

    /// <summary>
    /// Number format code, e.g. <c>#,##0.00</c> or <c>yyyy-mm-dd</c>.
    /// An empty string means <c>General</c>. Codes matching builtin Excel formats
    /// are written with their builtin ids.
    /// </summary>
    public string NumberFormat { get; init; } = "";

    /// <summary>
    /// Index of the named cell style (§18.8.8) this format descends from — the <c>xfId</c> of a
    /// loaded <c>xf</c>. It takes part in record equality on purpose: two visually identical
    /// formats attached to different named styles are different formats, and collapsing them
    /// would rebind cells to the wrong name. Styles built from code stay at 0, Normal.
    /// </summary>
    internal int StyleXfId { get; init; }
}

/// <summary>
/// Handle of a style registered in a workbook or streaming writer. Obtain via
/// <see cref="XlsxWriter.RegisterStyle"/>; the default value is the workbook default style.
/// </summary>
public readonly record struct StyleId
{
    internal StyleId(int index) => Index = index;

    internal int Index { get; }

    /// <summary>The workbook default style.</summary>
    public static StyleId Default => default;
}
