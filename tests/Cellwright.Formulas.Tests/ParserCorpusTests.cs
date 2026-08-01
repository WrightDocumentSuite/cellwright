using Cellwright.Formulas.Parsing;
using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Robustness of the parser on formulas nobody here wrote: every formula found in the sample
/// workbooks must either parse or fail with a proper parse exception, and rendering a parsed
/// program must produce text that parses back to the same program.
/// </summary>
public sealed class ParserCorpusTests
{
    [Fact]
    public async Task EveryFormulaInTheCorpus_ParsesOrFailsCleanly()
    {
        IReadOnlyList<string> files = OracleCorpus.Files(400);
        Assert.SkipWhen(files.Count == 0, "No sample workbooks were found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        int parsed = 0;
        int rejected = 0;
        var rejects = new List<string>();

        foreach (string file in files)
        {
            Workbook book;
            try
            {
                book = await Workbook.LoadAsync(file, new LoadOptions(), TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (OracleCorpus.IsUnreadable(exception))
            {
                continue;
            }

            foreach (Worksheet sheet in book.Sheets)
            {
                foreach ((int _, int _, string? formula) in sheet.Formulas.EnumerateOccupied())
                {
                    if (formula is null)
                        continue;

                    if (TryParse(formula, book))
                    {
                        parsed++;
                    }
                    else if (rejected++ < 20)
                    {
                        rejects.Add(formula);
                    }
                }
            }
        }

        TestContext.Current.TestOutputHelper?.WriteLine($"Parsed {parsed} formula(s), rejected {rejected}.");
        foreach (string reject in rejects)
            TestContext.Current.TestOutputHelper?.WriteLine($"  rejected: {reject}");

        Assert.True(parsed > 100, $"Only {parsed} formulas were found; the corpus looks wrong.");

        // The rejects are structured table references and external workbook links, both listed
        // as out of scope; the ratio guards against the parser regressing on ordinary formulas.
        double acceptance = (double)parsed / (parsed + rejected);
        Assert.True(acceptance > 0.85, $"Only {acceptance:P1} of real formulas parsed.");
    }

    [Fact]
    public async Task RenderingAParsedFormula_ProducesTheSameProgramAgain()
    {
        IReadOnlyList<string> files = OracleCorpus.Files(400);
        Assert.SkipWhen(files.Count == 0, "No sample workbooks were found. These checks read the sample workbooks that ship with ClosedXML, EPPlus, NPOI and ExcelDataReader; unpack those repositories beside this one to run them.");

        int checkedFormulas = 0;
        foreach (string file in files)
        {
            Workbook book;
            try
            {
                book = await Workbook.LoadAsync(file, new LoadOptions(), TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (OracleCorpus.IsUnreadable(exception))
            {
                continue;
            }

            foreach (Worksheet sheet in book.Sheets)
            {
                foreach ((int _, int _, string? formula) in sheet.Formulas.EnumerateOccupied())
                {
                    if (formula is null)
                        continue;

                    FormulaProgram? first = TryCompile(formula, book);
                    if (first is null)
                        continue;

                    string rendered = FormulaRenderer.Render(first, book);
                    FormulaProgram? second = TryCompile(rendered, book);
                    Assert.True(second is not null, $"'{formula}' rendered as '{rendered}', which no longer parses.");
                    Assert.True(
                        rendered == FormulaRenderer.Render(second, book),
                        $"'{formula}' rendered as '{rendered}' but re-rendered differently.");
                    checkedFormulas++;
                }
            }
        }

        Assert.True(checkedFormulas > 100, $"Only {checkedFormulas} formulas were round-tripped.");
    }

    [Fact]
    public void RandomInput_NeverEscapesAsAnUnexpectedException()
    {
        const string Alphabet = "AB1$:,;()+-*/^&<>=%\"'{} \t#!?_.\\ЖabZ";
        var random = new Random(20260725);
        Span<char> buffer = stackalloc char[40];

        for (int i = 0; i < 20_000; i++)
        {
            int length = random.Next(1, buffer.Length);
            for (int c = 0; c < length; c++)
                buffer[c] = Alphabet[random.Next(Alphabet.Length)];

            string formula = new(buffer[..length]);
            try
            {
                FormulaParser.Parse(formula);
            }
            catch (FormulaParseException)
            {
                // The only failure mode the parser is allowed to have.
            }
            catch (Exception exception)
            {
                Assert.Fail($"Parsing '{formula}' threw {exception.GetType().Name}: {exception.Message}");
            }
        }
    }

    private static bool TryParse(string formula, Workbook book) => TryCompile(formula, book) is not null;

    private static FormulaProgram? TryCompile(string formula, Workbook book)
    {
        try
        {
            return FormulaParser.Parse(formula, book);
        }
        catch (FormulaParseException)
        {
            return null;
        }
    }
}
