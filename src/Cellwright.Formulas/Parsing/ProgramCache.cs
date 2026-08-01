namespace Cellwright.Formulas.Parsing;

/// <summary>
/// Caches compiled programs by formula text. Spreadsheets repeat the same formula down whole
/// columns, so this turns "parse once per cell" into "parse once per distinct formula".
/// </summary>
internal sealed class ProgramCache(Workbook? workbook)
{
    private readonly Dictionary<string, FormulaProgram> _programs = new(StringComparer.Ordinal);

    /// <summary>Number of distinct programs held.</summary>
    public int Count => _programs.Count;

    /// <summary>
    /// Compiles the text, reusing an earlier result when possible. Parse failures are cached
    /// as well so a broken formula is not re-parsed on every recalculation.
    /// </summary>
    public FormulaProgram? GetOrParse(string text, out FormulaParseException? failure)
    {
        failure = null;
        if (_programs.TryGetValue(text, out FormulaProgram? cached))
            return cached;

        try
        {
            FormulaProgram program = FormulaParser.Parse(text, workbook);
            _programs[text] = program;
            return program;
        }
        catch (FormulaParseException exception)
        {
            failure = exception;
            return null;
        }
    }

    /// <summary>Drops everything; used when the workbook structure changes under the engine.</summary>
    public void Clear() => _programs.Clear();
}
