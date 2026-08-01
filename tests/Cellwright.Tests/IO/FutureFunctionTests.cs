using System.IO.Compression;
using System.Text;
using Xunit;

namespace Cellwright.Tests.IO;

/// <summary>
/// The <c>_xlfn.</c> prefix of MS-XLSX 2.2.3. Excel refuses to recognise a function the
/// standard never listed unless the file spells it with the prefix, so the writer has to add
/// it and the reader has to take it off again.
/// </summary>
public sealed class FutureFunctionTests
{
    [Theory]
    [InlineData("XLOOKUP(A1,B:B,C:C)", "_xlfn.XLOOKUP(A1,B:B,C:C)")]
    [InlineData("IFS(A1>0,1,TRUE,2)", "_xlfn.IFS(A1>0,1,TRUE,2)")]
    [InlineData("LET(x,3,x^2)", "_xlfn.LET(x,3,x^2)")]
    [InlineData("SUM(A1:A9)", "SUM(A1:A9)")]
    [InlineData("SUM(A1:A9)+CONCAT(B1,B2)", "SUM(A1:A9)+_xlfn.CONCAT(B1,B2)")]
    [InlineData("IF(ISFORMULA(A1),TEXTJOIN(\",\",1,B:B),\"\")", "IF(_xlfn.ISFORMULA(A1),_xlfn.TEXTJOIN(\",\",1,B:B),\"\")")]
    [InlineData("SORT(A1:A9)", "_xlfn._xlws.SORT(A1:A9)")]
    public void QualifyingAddsThePrefixWhereTheFormatNeedsIt(string bare, string qualified)
    {
        Assert.Equal(qualified, FutureFunctions.QualifyAll(bare));
        Assert.Equal(bare, FutureFunctions.UnqualifyAll(qualified));
    }

    [Fact]
    public void QualifyingIsIdempotent()
    {
        const string Qualified = "_xlfn.XLOOKUP(A1,B:B,C:C)";
        Assert.Equal(Qualified, FutureFunctions.QualifyAll(Qualified));
    }

    /// <summary>
    /// A prefixed name this library has never heard of belongs to a newer Excel. Dropping the
    /// prefix would break the formula, so the name travels exactly as it arrived.
    /// </summary>
    [Fact]
    public void AnUnknownPrefixedFunctionKeepsItsPrefix()
    {
        const string Formula = "_xlfn.SOMETHINGNEW(A1)";
        Assert.Equal(Formula, FutureFunctions.UnqualifyAll(Formula));
        Assert.Equal(Formula, FutureFunctions.QualifyAll(Formula));
    }

    /// <summary>A function name is a name only outside a quoted run.</summary>
    [Theory]
    [InlineData("\"IFS(\"&A1")]
    [InlineData("'My IFS(' & A1")]
    [InlineData("A1&\"XLOOKUP(x)\"")]
    public void TextThatLooksLikeACallIsLeftAlone(string formula) =>
        Assert.Equal(formula, FutureFunctions.QualifyAll(formula));

    /// <summary>A name that merely starts with a future function's letters is not one.</summary>
    [Theory]
    [InlineData("XLOOKUPS(A1)")]
    [InlineData("MYIFS(A1)")]
    [InlineData("IFS")]
    [InlineData("A1+IFS")]
    public void OnlyAWholeNameFollowedByABracketCounts(string formula) =>
        Assert.Equal(formula, FutureFunctions.QualifyAll(formula));

    [Fact]
    public async Task TheWrittenFileCarriesThePrefix()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("XLOOKUP(1,B:B,C:C)");
        sheet.Cell("A2").SetFormula("SUM(B1:B2)");

        string xml = await SaveAndReadSheetAsync(book);

        Assert.Contains("<f>_xlfn.XLOOKUP(1,B:B,C:C)</f>", xml, StringComparison.Ordinal);
        Assert.Contains("<f>SUM(B1:B2)</f>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheModelHoldsTheNameAPersonWouldType()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetFormula("XLOOKUP(1,B:B,C:C)");

        using var buffer = new MemoryStream();
        await book.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        Workbook reloaded = await Workbook.LoadAsync(
            buffer, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        Assert.Equal("XLOOKUP(1,B:B,C:C)", reloaded.Sheets[0].Cell("A1").Formula);
    }

    [Fact]
    public async Task DefinedNamesAndValidationsAreSpelledTheSameWay()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet["A1"] = 1;
        book.Names.Add("Recent", "S!$A$1:INDEX(S!$A:$A,DAYS(TODAY(),S!$A$1))");
        sheet.Validations.Add(DataValidation.Custom("B1:B9", "ISFORMULA(B1)"));

        string sheetXml = await SaveAndReadSheetAsync(book);
        Assert.Contains("_xlfn.ISFORMULA(B1)", sheetXml, StringComparison.Ordinal);

        string workbookXml = await SaveAndReadPartAsync(book, "xl/workbook.xml");
        Assert.Contains("_xlfn.DAYS(", workbookXml, StringComparison.Ordinal);
    }

    private static Task<string> SaveAndReadSheetAsync(Workbook book) =>
        SaveAndReadPartAsync(book, "xl/worksheets/sheet1.xml");

    private static async Task<string> SaveAndReadPartAsync(Workbook book, string partName)
    {
        using var buffer = new MemoryStream();
        await book.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        using var zip = new ZipArchive(buffer, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(partName)!.Open(), Encoding.UTF8);
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }
}
