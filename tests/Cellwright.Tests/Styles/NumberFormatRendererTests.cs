using System.Globalization;
using Cellwright.Styles;
using Xunit;

namespace Cellwright.Tests.Styles;

/// <summary>
/// The number format renderer, checked against what Excel puts in the cell. Every case here was
/// read off a real sheet rather than derived from the format grammar.
/// </summary>
public sealed class NumberFormatRendererTests
{
    private static string Format(CellValue value, string code) =>
        NumberFormatRenderer.Format(value, code, CultureInfo.InvariantCulture);

    [Theory]
    [InlineData(1234.567, "0", "1235")]
    [InlineData(1234.567, "0.00", "1234.57")]
    [InlineData(1234.567, "#,##0.00", "1,234.57")]
    [InlineData(1234567.0, "#,##0", "1,234,567")]
    [InlineData(0.5, "0.00", "0.50")]
    [InlineData(0.5, ".00", ".50")]
    [InlineData(0.5, "#.##", ".5")]
    [InlineData(-1234.5, "#,##0.00", "-1,234.50")]
    [InlineData(0.0, "0.00", "0.00")]
    [InlineData(12.0, "0000", "0012")]
    [InlineData(0.256, "0%", "26%")]
    [InlineData(0.256, "0.0%", "25.6%")]
    [InlineData(1234567.0, "#,##0,", "1,235")]
    [InlineData(1234567.0, "0.0,,", "1.2")]
    public void Numbers(double value, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(value), code));
    }

    [Theory]
    [InlineData(-1234.5, "#,##0.00;(#,##0.00)", "(1,234.50)")]
    [InlineData(1234.5, "#,##0.00;(#,##0.00)", "1,234.50")]
    [InlineData(0.0, "0.00;(0.00);\"zero\"", "zero")]
    [InlineData(-5.0, "0;[Red]-0", "-5")]
    public void SectionsPickThemselvesBySign(double value, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(value), code));
    }

    [Theory]
    [InlineData(12345.0, "0.00E+00", "1.23E+04")]
    [InlineData(0.00012, "0.00E+00", "1.20E-04")]
    [InlineData(12345.0, "0.0E+0", "1.2E+4")]
    public void ScientificNotation(double value, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(value), code));
    }

    /// <summary>
    /// Fractions. The room the format leaves the denominator is the whole instruction: one digit
    /// can only reach halves and quarters and thirds, three digits get close to anything.
    /// </summary>
    [Theory]
    [InlineData(2.5, "# ?/?", "2 1/2")]
    [InlineData(2.25, "# ?/?", "2 1/4")]
    [InlineData(0.75, "# ?/?", " 3/4")]
    [InlineData(-2.5, "# ?/?", "-2 1/2")]
    [InlineData(2.75, "# ??/16", "2 12/16")]
    [InlineData(0.3125, "# ??/16", "  5/16")]
    [InlineData(1.0 / 3.0, "?/?", "1/3")]
    [InlineData(2.0 / 3.0, "# ?/?", " 2/3")]
    [InlineData(3.14159265358979, "# ???/???", "3  16/113")]
    [InlineData(3.14159265358979, "?/???", "355/113")]
    public void Fractions(double value, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(value), code));
    }

    /// <summary>
    /// A whole number has no ratio to show, and a rounding that fills the denominator carries
    /// into the whole part rather than printing a ratio of one.
    /// </summary>
    [Fact]
    public void AWholeNumberLeavesTheRatioBlank()
    {
        Assert.Equal("5    ", Format(CellValue.Number(5), "# ?/?"));
        Assert.Equal("5 ", Format(CellValue.Number(5), "# #/#"));

        // A single digit cannot get closer to 0.99 than one whole, so the whole part takes it.
        Assert.Equal("1    ", Format(CellValue.Number(0.99), "# ?/?"));

        // Nine tenths, on the other hand, is nearer eight ninths than it is to one.
        Assert.Equal(" 8/9", Format(CellValue.Number(0.9), "# ?/?"));
    }

    [Theory]
    [InlineData(1234.5, "$#,##0.00", "$1,234.50")]
    [InlineData(1234.5, "#,##0.00\" руб.\"", "1,234.50 руб.")]
    [InlineData(5.0, "0 \\k\\g", "5 kg")]
    [InlineData(5.0, "_-0_-", "5")]
    public void LiteralsAndPadding(double value, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(value), code));
    }

    [Theory]
    [InlineData("2026-07-25", "yyyy-mm-dd", "2026-07-25")]
    [InlineData("2026-07-25", "d/m/yyyy", "25/7/2026")]
    [InlineData("2026-07-25", "dd-mmm-yy", "25-Jul-26")]
    [InlineData("2026-07-25", "mmmm yyyy", "July 2026")]
    [InlineData("2026-07-25", "mmmmm", "J")]
    [InlineData("2026-07-25", "dddd", "Saturday")]
    [InlineData("2026-07-25", "ddd", "Sat")]
    [InlineData("2026-07-25 14:05:09", "h:mm", "14:05")]
    [InlineData("2026-07-25 14:05:09", "hh:mm:ss", "14:05:09")]
    [InlineData("2026-07-25 14:05:09", "h:mm AM/PM", "2:05 PM")]
    [InlineData("2026-07-25 09:05:09", "h:mm AM/PM", "9:05 AM")]
    [InlineData("2026-07-25 09:05:09", "h:mm A/P", "9:05 A")]
    [InlineData("2026-07-25 14:05:09", "yyyy-mm-dd hh:mm", "2026-07-25 14:05")]
    [InlineData("2026-01-05", "m/d/yy", "1/5/26")]
    public void DatesAndTimes(string moment, string code, string expected)
    {
        var value = CellValue.DateTime(DateTime.Parse(moment, CultureInfo.InvariantCulture));

        Assert.Equal(expected, Format(value, code));
    }

    [Theory]
    [InlineData(1.5, "[h]:mm", "36:00")]
    [InlineData(1.5, "[m]", "2160")]
    [InlineData(0.5, "[h]:mm:ss", "12:00:00")]
    public void ElapsedTimesCountPastMidnight(double days, string code, string expected)
    {
        Assert.Equal(expected, Format(CellValue.Number(days), code));
    }

    [Fact]
    public void TextGoesThroughTheFourthSectionOrUntouched()
    {
        Assert.Equal("hello", Format(CellValue.Text("hello"), "0.00"));
        Assert.Equal("<hello>", Format(CellValue.Text("hello"), "0.00;-0.00;0.00;\"<\"@\">\""));
        Assert.Equal("", Format(CellValue.Blank, "0.00"));
    }

    [Fact]
    public void GeneralAndAnEmptyCodeShowTheValueAsIs()
    {
        Assert.Equal("1234.5", Format(CellValue.Number(1234.5), "General"));
        Assert.Equal("1234.5", Format(CellValue.Number(1234.5), ""));
    }

    [Fact]
    public void CultureDecidesTheSeparators()
    {
        var german = CultureInfo.GetCultureInfo("de-DE");

        Assert.Equal("1.234,57", NumberFormatRenderer.Format(CellValue.Number(1234.567), "#,##0.00", german));
        Assert.Equal("Juli", NumberFormatRenderer.Format(
            CellValue.DateTime(new DateTime(2026, 7, 25)), "mmmm", german));
    }
}
