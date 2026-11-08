using Cellwright.Styles;
using Xunit;

namespace Cellwright.Pdf.Tests;

/// <summary>
/// Sizing columns and rows to their contents: what widens a column, what does not, and what the
/// limits do.
/// </summary>
public sealed class AutoFitTests
{
    private static Worksheet Sheet(Workbook book) => book.Sheets.Add("Sheet1");

    [Fact]
    public void LongerTextGetsAWiderColumn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "x";
        sheet["B1"] = "a considerably longer piece of text";

        sheet.AutoFitColumns();

        Assert.NotNull(sheet.Column("A").Width);
        Assert.NotNull(sheet.Column("B").Width);
        Assert.True(sheet.Column("B").Width > sheet.Column("A").Width);
    }

    [Fact]
    public void ANumberIsMeasuredAtTheWidthItsFormatPrints()
    {
        // 0.5 under this format draws as "50.00%", which is five characters wider than the
        // number itself: measuring the raw value would size the column to the wrong string.
        Workbook plain = Workbook.Create();
        Worksheet plainSheet = Sheet(plain);
        plainSheet["A1"] = 0.5;
        plainSheet.AutoFitColumns();

        Workbook formatted = Workbook.Create();
        Worksheet formattedSheet = Sheet(formatted);
        formattedSheet["A1"] = 0.5;
        formattedSheet.Cell("A1").SetStyle(CellStyle.Default with { NumberFormat = "0.00%" });
        formattedSheet.AutoFitColumns();

        Assert.True(formattedSheet.Column("A").Width > plainSheet.Column("A").Width);
    }

    [Fact]
    public void BoldTextAsksForMoreRoomThanPlain()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        const string Text = "Quarterly revenue";
        sheet["A1"] = Text;
        sheet["B1"] = Text;
        sheet.Cell("B1").SetStyle(CellStyle.Default with { Font = FontStyle.Default with { Bold = true } });

        sheet.AutoFitColumns();

        Assert.True(sheet.Column("B").Width > sheet.Column("A").Width);
    }

    [Fact]
    public void TheWidthCeilingIsRespected()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = new string('m', 400);

        sheet.AutoFitColumns(new AutoFitOptions { MaxWidth = 20 });

        Assert.Equal(20, sheet.Column("A").Width);
    }

    [Fact]
    public void TheWidthFloorIsRespected()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "i";

        sheet.AutoFitColumns(new AutoFitOptions { MinWidth = 12 });

        Assert.Equal(12, sheet.Column("A").Width);
    }

    [Fact]
    public void NothingGrowsWiderThanTheFormatAllows()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = new string('W', 5000);

        sheet.AutoFitColumns();

        Assert.True(sheet.Column("A").Width <= SheetLimits.MaxColumnWidth);
    }

    [Fact]
    public void AWrappedCellSizesItsRowRatherThanItsColumn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "a sentence long enough that it has to be broken over several lines";
        sheet.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        sheet.Column("A").SetWidth(12);

        sheet.AutoFit();

        // The column keeps the width the author chose, and the row grew to hold the wrapping.
        Assert.Equal(12, sheet.Column("A").Width);
        Assert.NotNull(sheet.Row(1).Height);
        Assert.True(sheet.Row(1).Height > 15);
    }

    [Fact]
    public void AMergedCellDoesNotSizeAnyColumn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "short";
        sheet["A2"] = "a merged banner that runs across both columns and is very long indeed";
        sheet.Merge(RangeRef.Parse("A2:B2"));

        sheet.AutoFitColumns();

        // Only A1 had a say, so the column is sized to "short" rather than to the banner.
        double? fitted = sheet.Column("A").Width;
        Assert.NotNull(fitted);

        Workbook alone = Workbook.Create();
        Worksheet aloneSheet = Sheet(alone);
        aloneSheet["A1"] = "short";
        aloneSheet.AutoFitColumns();

        Assert.Equal(aloneSheet.Column("A").Width, fitted);
    }

    [Fact]
    public void AMergedCellCanOptInToSizingItsRow()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "A long merged field that has to wrap across several lines to remain visible";
        sheet.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        sheet.Column("A").SetWidth(7);
        sheet.Column("B").SetWidth(7);
        sheet.Merge(RangeRef.Parse("A1:B1"));

        sheet.AutoFitRows();
        Assert.Null(sheet.Row(1).Height);

        sheet.AutoFitRows(new AutoFitOptions { IncludeMergedCells = true });

        Assert.NotNull(sheet.Row(1).Height);
        Assert.True(sheet.Row(1).Height > 15);
    }

    [Fact]
    public void AHorizontalMergeIsMeasuredAtItsCombinedWidth()
    {
        const string Text = "A long merged field whose wrapping changes with the width of the merged range";

        Workbook narrowBook = Workbook.Create();
        Worksheet narrow = Sheet(narrowBook);
        narrow["A1"] = Text;
        narrow.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        narrow.Column("A").SetWidth(7);
        narrow.Column("B").SetWidth(7);
        narrow.Merge(RangeRef.Parse("A1:B1"));
        narrow.AutoFitRows(new AutoFitOptions { IncludeMergedCells = true });

        Workbook wideBook = Workbook.Create();
        Worksheet wide = Sheet(wideBook);
        wide["A1"] = Text;
        wide.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        for (int column = 1; column <= 4; column++)
            wide.Column(column).SetWidth(7);
        wide.Merge(RangeRef.Parse("A1:D1"));
        wide.AutoFitRows(new AutoFitOptions { IncludeMergedCells = true });

        Assert.True(narrow.Row(1).Height > wide.Row(1).Height);
    }

    [Fact]
    public void AMultiRowMergeGrowsOnlyWhenItsCombinedHeightIsTooShort()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = new string('m', 180);
        sheet.Cell("A1").SetStyle(
            CellStyle.Default with { Alignment = Alignment.Default with { WrapText = true } });
        sheet.Column("A").SetWidth(6);
        sheet.Merge(RangeRef.Parse("A1:A2"));

        sheet.AutoFitRows(new AutoFitOptions { IncludeMergedCells = true });

        double total = (sheet.Row(1).Height ?? 15) + (sheet.Row(2).Height ?? 15);
        Assert.True(total > 30);
    }

    [Fact]
    public void AHiddenRowDoesNotSizeAnything()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "short";
        sheet["A2"] = "a hidden row holding something far longer than the visible one";
        sheet.Row(2).Hide();

        sheet.AutoFitColumns();
        double? withHidden = sheet.Column("A").Width;

        sheet.AutoFitColumns(new AutoFitOptions { IncludeHiddenRows = true });
        double? withoutHidden = sheet.Column("A").Width;

        Assert.True(withoutHidden > withHidden);
    }

    [Fact]
    public void OnlyTheSampledRowsAreMeasured()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "short";
        sheet["A50"] = "a row far below the sample that holds something much longer";

        sheet.AutoFitColumns(new AutoFitOptions { SampleRows = 10 });
        double? sampled = sheet.Column("A").Width;

        sheet.AutoFitColumns();
        double? everything = sheet.Column("A").Width;

        Assert.True(everything > sampled);
    }

    [Fact]
    public void AFittedColumnIsWideEnoughForItsTextWhenPrinted()
    {
        // The point of measuring in the printing package: what was fitted must not be cut.
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        const string Text = "Weighted average cost of capital";
        sheet["A1"] = Text;

        sheet.AutoFitColumns();

        using Rendered rendered = Rendered.Of(book);
        Assert.Contains(Text, rendered.Text(0), StringComparison.Ordinal);
    }

    [Fact]
    public void FittingTheWorkbookFitsEverySheet()
    {
        Workbook book = Workbook.Create();
        Worksheet first = book.Sheets.Add("First");
        Worksheet second = book.Sheets.Add("Second");
        first["A1"] = "a piece of text on the first sheet";
        second["A1"] = "a piece of text on the second sheet";

        book.AutoFit();

        Assert.NotNull(first.Column("A").Width);
        Assert.NotNull(second.Column("A").Width);
    }

    [Fact]
    public void OneColumnCanBeFittedOnItsOwn()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);
        sheet["A1"] = "a piece of text in the first column";
        sheet["B1"] = "a piece of text in the second column";

        sheet.Column("B").AutoFit();

        Assert.Null(sheet.Column("A").Width);
        Assert.NotNull(sheet.Column("B").Width);
    }

    [Fact]
    public void AnEmptySheetIsLeftAlone()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Sheet(book);

        sheet.AutoFit();

        Assert.Null(sheet.Column("A").Width);
        Assert.Null(sheet.Row(1).Height);
    }
}
