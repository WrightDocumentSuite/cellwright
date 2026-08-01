using System.Collections;
using System.Text;

namespace Cellwright;

/// <summary>
/// One stretch of a cell's text drawn in its own character formatting.
/// </summary>
/// <remarks>
/// A run without a <see cref="Font"/> is drawn in the cell's own font, which is what the format
/// means by a run with no <c>rPr</c>. A run that names one names it whole rather than as a set of
/// overrides: the format has no way to say "the cell's font but bold", so a bold run carries the
/// family, the size and the colour too. Deriving from the cell's font — <c>cell.StyleValue.Font
/// with { Bold = true }</c> — is how you say that and mean it.
/// </remarks>
/// <param name="Text">The characters of the run.</param>
/// <param name="Font">How they are drawn, or <see langword="null"/> to follow the cell's font.</param>
public sealed record RichTextRun(string Text, FontStyle? Font = null);

/// <summary>
/// A cell's text with character formatting inside it: several runs, each with its own font.
/// </summary>
/// <remarks>
/// Immutable, like the styles. <see cref="Append(string, FontStyle)"/> returns a new value rather
/// than changing this one, which is cheap here because a formatted string is a handful of runs
/// rather than a document.
/// </remarks>
/// <example>
/// <code>
/// FontStyle body = sheet.Cell("A1").StyleValue.Font;
/// sheet.Cell("A1").SetRichText(RichText.Empty
///     .Append("Total: ")
///     .Append("42", body with { Bold = true }));
/// </code>
/// </example>
public sealed class RichText : IReadOnlyList<RichTextRun>
{
    private readonly RichTextRun[] _runs;
    private string? _text;

    /// <summary>A text with no runs at all.</summary>
    public static RichText Empty { get; } = new([]);

    /// <summary>Builds a formatted text from its runs.</summary>
    /// <param name="runs">The runs, in reading order.</param>
    public RichText(params IEnumerable<RichTextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);
        _runs = [.. runs];

        foreach (RichTextRun run in _runs)
            ArgumentNullException.ThrowIfNull(run);
    }

    /// <inheritdoc/>
    public int Count => _runs.Length;

    /// <inheritdoc/>
    public RichTextRun this[int index] => _runs[index];

    /// <summary>The visible text, which is every run's text end to end.</summary>
    public string Text => _text ??= Join(_runs);

    /// <summary>Whether any run carries formatting of its own.</summary>
    /// <remarks>
    /// A text whose runs all follow the cell's font is an ordinary string wearing several coats
    /// of the same colour: it is written as one, because nothing would be lost.
    /// </remarks>
    public bool IsFormatted
    {
        get
        {
            foreach (RichTextRun run in _runs)
            {
                if (run.Font is not null)
                    return true;
            }

            return false;
        }
    }

    /// <summary>A copy of this text with one more run at the end.</summary>
    /// <param name="text">The characters of the run.</param>
    /// <param name="font">How they are drawn.</param>
    public RichText Append(string text, FontStyle font)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(font);
        return new RichText([.. _runs, new RichTextRun(text, font)]);
    }

    /// <summary>A copy of this text with one more run at the end, drawn in the cell's own font.</summary>
    /// <param name="text">The characters of the run.</param>
    public RichText Append(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new RichText([.. _runs, new RichTextRun(text)]);
    }

    /// <summary>A plain string as a single unformatted run.</summary>
    /// <param name="text">The text.</param>
    public static RichText FromPlain(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new RichText(new RichTextRun(text));
    }

    /// <inheritdoc/>
    public IEnumerator<RichTextRun> GetEnumerator() => ((IReadOnlyList<RichTextRun>)_runs).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc/>
    public override string ToString() => Text;

    private static string Join(RichTextRun[] runs)
    {
        if (runs.Length == 0)
            return string.Empty;
        if (runs.Length == 1)
            return runs[0].Text;

        var builder = new StringBuilder();
        foreach (RichTextRun run in runs)
            builder.Append(run.Text);

        return builder.ToString();
    }
}
