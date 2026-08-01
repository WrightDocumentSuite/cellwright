using System.Text;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Pins down the shapes of an encoded formula that a round trip through this library cannot
/// check. Writing and reading with the same misreading agrees with itself, so each of these was
/// found by decoding a real workbook alongside another library and is anchored on that workbook.
/// </summary>
public sealed class FormulaDecodingTests
{
    static FormulaDecodingTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    private static string? Find(string name) =>
        XlsCorpus.Files.FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    private static async Task<Workbook> WorkbookAsync(string file)
    {
        string? path = Find(file);
        Assert.SkipWhen(path is null, $"{file} is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        return await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);
    }

    /// <summary>The first formula anywhere in the workbook that mentions <paramref name="marker"/>.</summary>
    private static string FormulaMentioning(Workbook workbook, string marker)
    {
        foreach (Worksheet sheet in workbook.Sheets)
        {
            foreach ((int _, int _, string? formula) in sheet.Formulas.EnumerateOccupied())
            {
                if (formula is not null && formula.Contains(marker, StringComparison.Ordinal))
                    return formula;
            }
        }

        Assert.Fail($"No formula mentions '{marker}'.");
        return "";
    }

    /// <summary>
    /// A function borrowed from an add-in is not named in the formula at all: the token points
    /// at an ExternName record beside the workbook's link table. Without reading those, every
    /// call to one of the Analysis ToolPak functions decodes as a reference to nothing.
    /// </summary>
    [Fact]
    public async Task AFunctionBorrowedFromAnAddInKeepsItsName()
    {
        Workbook workbook = await WorkbookAsync("15228.xls");

        string formula = FormulaMentioning(workbook, "TrendFacDump");
        Assert.Contains("EOMONTH(ValDat,-1)", formula, StringComparison.Ordinal);
        Assert.DoesNotContain("#REF!", formula, StringComparison.Ordinal);
    }

    /// <summary>
    /// A reference into another workbook counts that workbook's tabs, not this one's. Reading
    /// the index against the local list names a sheet that exists here and reads its data, so
    /// the formula comes out plausible and wrong rather than obviously broken.
    /// </summary>
    [Fact]
    public async Task AReferenceIntoAnotherWorkbookNamesThatWorkbook()
    {
        Workbook workbook = await WorkbookAsync("19599-1.xls");

        string formula = FormulaMentioning(workbook, "VLOOKUP");
        Assert.Contains("0231324V1-1.xls]Sheet1", formula, StringComparison.Ordinal);
    }

    /// <summary>
    /// Whole columns are the only way the old format can say "this column", and it says it as an
    /// area spanning every row. Spelled back out as $D$1:$I$65536 the range is the same, but the
    /// shape is lost and with it the clipping that keeps a lookup from walking 65 536 rows.
    /// </summary>
    [Fact]
    public async Task AWholeColumnRangeKeepsItsShape()
    {
        Workbook workbook = await WorkbookAsync("19599-1.xls");

        string formula = FormulaMentioning(workbook, "VLOOKUP");
        Assert.Contains("!$D:$I", formula, StringComparison.Ordinal);
        Assert.DoesNotContain("65536", formula, StringComparison.Ordinal);
    }

    /// <summary>
    /// A shared formula states its references as offsets, and an offset that runs off an edge
    /// wraps to the far side rather than stopping at it.
    /// </summary>
    [Fact]
    public async Task ARelativeReferencePastAnEdgeWrapsRound()
    {
        Workbook workbook = await WorkbookAsync("OldIssue11573_BlankValues.xls");

        Assert.Equal("VLOOKUP(C13,IR:IS,2,FALSE)", workbook.Sheets[0].Formulas.Get(13, 2));
    }
}
