using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Resolving a sheet-qualified range from text: the spelling a chart series or a print area
/// holds, back to the sheet and rectangle it names.
/// </summary>
public sealed class SheetRangeResolverTests
{
    private static Workbook Book(out Worksheet sheet, out Worksheet second)
    {
        var book = Workbook.Create();
        sheet = book.Sheets.Add("Данные");
        second = book.Sheets.Add("Q1 2026");
        return book;
    }

    [Fact]
    public void AQualifiedRangeNamesItsSheetAndRectangle()
    {
        Workbook book = Book(out Worksheet sheet, out _);

        (Worksheet resolved, RangeRef range) = SheetRangeResolver.Resolve(book, "Данные!$B$2:$B$10")!.Value;

        Assert.Same(sheet, resolved);
        Assert.Equal(RangeRef.Parse("B2:B10"), range);
    }

    [Fact]
    public void AQuotedSheetNameWithASpaceResolves()
    {
        Workbook book = Book(out _, out Worksheet second);

        (Worksheet resolved, RangeRef range) = SheetRangeResolver.Resolve(book, "'Q1 2026'!A1:C5")!.Value;

        Assert.Same(second, resolved);
        Assert.Equal(RangeRef.Parse("A1:C5"), range);
    }

    [Fact]
    public void ASingleCellIsARectangleOfOne()
    {
        Workbook book = Book(out Worksheet sheet, out _);

        (Worksheet resolved, RangeRef range) = SheetRangeResolver.Resolve(book, "Данные!C7")!.Value;

        Assert.Same(sheet, resolved);
        Assert.Equal(RangeRef.Parse("C7"), range);
    }

    [Fact]
    public void ALeadingEqualsSignIsAllowed()
    {
        Workbook book = Book(out Worksheet sheet, out _);

        Assert.NotNull(SheetRangeResolver.Resolve(book, "=Данные!A1:A5"));
    }

    [Fact]
    public void ADefinedNameStandsForItsReference()
    {
        Workbook book = Book(out Worksheet sheet, out _);
        book.Names.Add("Продажи", "Данные!$A$2:$A$99");

        (Worksheet resolved, RangeRef range) = SheetRangeResolver.Resolve(book, "Продажи")!.Value;

        Assert.Same(sheet, resolved);
        Assert.Equal(RangeRef.Parse("A2:A99"), range);
    }

    [Fact]
    public void AWholeColumnExpandsToTheGrid()
    {
        Workbook book = Book(out Worksheet sheet, out _);

        (_, RangeRef range) = SheetRangeResolver.Resolve(book, "Данные!$D:$F")!.Value;

        Assert.Equal(4, range.Start.Column);
        Assert.Equal(6, range.End.Column);
        Assert.Equal(1, range.Start.Row);
        Assert.Equal(SheetLimits.MaxRow, range.End.Row);
    }

    [Fact]
    public void ABackwardsRangeIsTheSameRectangle()
    {
        Workbook book = Book(out _, out _);

        (_, RangeRef range) = SheetRangeResolver.Resolve(book, "Данные!B10:A2")!.Value;

        Assert.Equal(RangeRef.Parse("A2:B10"), range);
    }

    [Theory]
    [InlineData("Никакой!A1")]           // Sheet that does not exist.
    [InlineData("Данные!A1+B2")]         // An expression, not a reference.
    [InlineData("SUM(Данные!A1:A5)")]    // A formula.
    [InlineData("НеИмя")]                // A name nobody defined.
    [InlineData("")]
    public void AnythingElseAnswersNullRatherThanAGuess(string text)
    {
        Workbook book = Book(out _, out _);

        Assert.Null(SheetRangeResolver.Resolve(book, text));
    }

    [Fact]
    public void ANameChainIsFollowedButNotForever()
    {
        Workbook book = Book(out Worksheet sheet, out _);
        book.Names.Add("Один", "Два");
        book.Names.Add("Два", "Данные!$A$1");
        book.Names.Add("Сам", "Сам");

        Assert.NotNull(SheetRangeResolver.Resolve(book, "Один"));
        Assert.Null(SheetRangeResolver.Resolve(book, "Сам"));
    }
}
