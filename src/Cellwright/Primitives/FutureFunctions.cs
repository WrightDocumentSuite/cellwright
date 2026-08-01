using System.Collections.Frozen;
using System.Text;

namespace Cellwright;

/// <summary>
/// The functions Excel added after ISO/IEC 29500 was written, and the prefixes the file format
/// spells them with.
/// </summary>
/// <remarks>
/// A function the standard does not list has to be written as <c>_xlfn.NAME</c> — Excel shows
/// <c>#NAME?</c> for a bare <c>XLOOKUP(</c> even though it implements the function perfectly
/// well. The three worksheet-only ones take a second prefix, <c>_xlfn._xlws.</c>
/// (MS-XLSX 2.2.3). The prefix is a spelling of the file format and nothing more, so it is
/// stripped when a formula is read and restored when one is written: what the model holds, and
/// what <see cref="Cell.Formula"/> returns, is the name a person would type.
/// </remarks>
public static class FutureFunctions
{
    /// <summary>The prefix an unlisted function is written with.</summary>
    public const string Prefix = "_xlfn.";

    /// <summary>The additional prefix of a function that only works on a worksheet.</summary>
    public const string WorksheetPrefix = "_xlws.";

    private static readonly FrozenSet<string> Names = BuildNames();

    private static readonly FrozenSet<string> WorksheetOnlyNames =
        new[] { "FILTER", "PY", "SORT" }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Whether a function has to carry a prefix when written to a file.</summary>
    public static bool NeedsPrefix(string name) => Names.Contains(name) || WorksheetOnlyNames.Contains(name);

    /// <summary>Returns the name as a file spells it, e.g. <c>XLOOKUP</c> to <c>_xlfn.XLOOKUP</c>.</summary>
    public static string Qualify(string name)
    {
        if (WorksheetOnlyNames.Contains(name))
            return Prefix + WorksheetPrefix + name;
        return Names.Contains(name) ? Prefix + name : name;
    }

    /// <summary>
    /// Rewrites every function call in a formula to the spelling a file uses. Returns the same
    /// string when nothing needed a prefix, which is the usual case.
    /// </summary>
    public static string? QualifyAll(string? formula) => Rewrite(formula, qualify: true);

    /// <summary>
    /// Rewrites every function call in a formula to its bare name, undoing
    /// <see cref="QualifyAll"/>. A prefixed name this table does not know is left alone: it
    /// belongs to a newer Excel than this, and dropping its prefix would break the file.
    /// </summary>
    public static string? UnqualifyAll(string? formula) => Rewrite(formula, qualify: false);

    /// <summary>
    /// Walks a formula's function names and respells them. Literals are stepped over: a
    /// quoted string may hold anything at all, including something that looks like a call.
    /// </summary>
    private static string? Rewrite(string? formula, bool qualify)
    {
        if (string.IsNullOrEmpty(formula) || !MightContainCall(formula))
            return formula;

        StringBuilder? builder = null;
        int copied = 0;
        int i = 0;

        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"' || c == '\'')
            {
                i = SkipQuoted(formula, i, c);
                continue;
            }

            if (!IsNameStart(c))
            {
                i++;
                continue;
            }

            int start = i;
            while (i < formula.Length && IsNamePart(formula[i]))
                i++;

            // Only a name immediately followed by an opening bracket is a call; anything else
            // is a defined name, a sheet name or a reference.
            if (i >= formula.Length || formula[i] != '(')
                continue;

            ReadOnlySpan<char> token = formula.AsSpan(start, i - start);
            string bare = StripPrefixes(token);

            // A name this table does not know travels exactly as written, whichever way we are
            // going: it may be a function from a newer Excel, and either adding a prefix or
            // taking one away would change what the formula means.
            if (!NeedsPrefix(bare))
                continue;

            string replacement = qualify ? Qualify(bare) : bare;
            if (token.SequenceEqual(replacement))
                continue;

            builder ??= new StringBuilder(formula.Length + 16);
            builder.Append(formula, copied, start - copied).Append(replacement);
            copied = i;
        }

        if (builder is null)
            return formula;

        return builder.Append(formula, copied, formula.Length - copied).ToString();
    }

    private static string StripPrefixes(ReadOnlySpan<char> token)
    {
        if (token.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
        {
            token = token[Prefix.Length..];
            if (token.StartsWith(WorksheetPrefix, StringComparison.OrdinalIgnoreCase))
                token = token[WorksheetPrefix.Length..];
        }

        return token.ToString();
    }

    /// <summary>A formula with no bracket has no call in it, and most formulas are references.</summary>
    private static bool MightContainCall(string formula) => formula.Contains('(', StringComparison.Ordinal);

    /// <summary>Steps over a quoted run, honouring the doubled quote that escapes one.</summary>
    private static int SkipQuoted(string text, int start, char quote)
    {
        int i = start + 1;
        while (i < text.Length)
        {
            if (text[i] != quote)
            {
                i++;
                continue;
            }

            if (i + 1 < text.Length && text[i + 1] == quote)
            {
                i += 2;
                continue;
            }

            return i + 1;
        }

        return i;
    }

    private static bool IsNameStart(char c) => char.IsAsciiLetter(c) || c == '_' || c == '\\' || c > 0x7F;

    private static bool IsNamePart(char c) => IsNameStart(c) || char.IsAsciiDigit(c) || c == '.';

    /// <summary>
    /// The functions of MS-XLSX 2.2.3 that carry the <c>_xlfn.</c> prefix. Four future
    /// functions are absent on purpose — <c>ECMA.CEILING</c>, <c>ISO.CEILING</c>,
    /// <c>NETWORKDAYS.INTL</c> and <c>WORKDAY.INTL</c> are in the standard's own list and are
    /// written bare.
    /// </summary>
    private static FrozenSet<string> BuildNames() => new[]
    {
        "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST",
        "BINOM.DIST.RANGE", "BINOM.INV", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR",
        "BYCOL", "BYROW", "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT",
        "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSECOLS", "CHOOSEROWS", "COMBINA",
        "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COPILOT", "COT", "COTH", "COVARIANCE.P",
        "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP", "ERF.PRECISE", "ERFC.PRECISE",
        "EXPAND", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
        "FIELDVALUE", "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS",
        "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT",
        "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV",
        "GAMMALN.PRECISE", "GAUSS", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMCOSH", "IMCOT",
        "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA", "ISOMITTED",
        "ISOWEEKNUM", "LAMBDA", "LET", "LOGNORM.DIST", "LOGNORM.INV", "LONGTEXT", "MAKEARRAY",
        "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT", "NEGBINOM.DIST",
        "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION",
        "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC",
        "PERMUTATIONA", "PHI", "POISSON.DIST", "PQSOURCE", "PYTHON_STR", "PYTHON_TYPE",
        "PYTHON_TYPENAME", "QUARTILE.EXC", "QUARTILE.INC", "QUERYSTRING", "RANDARRAY",
        "RANK.AVG", "RANK.EQ", "REDUCE", "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET",
        "SHEETS", "SINGLE", "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST", "T.DIST.2T",
        "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE",
        "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW", "UNICHAR", "UNICODE", "UNIQUE", "VAR.P",
        "VAR.S", "VSTACK", "WEBSERVICE", "WEIBULL.DIST", "WRAPCOLS", "WRAPROWS", "XLOOKUP",
        "XOR", "Z.TEST",
    }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);
}
