using Xunit;

namespace Cellwright.Tests.Primitives;

public sealed class ColumnNameTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16_384, "XFD")]
    public void Format_And_Parse_RoundTrip(int column, string name)
    {
        Assert.Equal(name, ColumnName.Format(column));
        Assert.Equal(column, ColumnName.Parse(name));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    [InlineData(16_385)]
    public void Format_OutOfRange_Throws(int column)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ColumnName.Format(column));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1")]
    [InlineData("XFE")]
    [InlineData("*")]
    public void Parse_Invalid_Throws(string name)
    {
        Assert.Throws<FormatException>(() => ColumnName.Parse(name));
    }

    [Fact]
    public void ExcelErrorText_RoundTrips()
    {
        foreach (ExcelError error in Enum.GetValues<ExcelError>())
        {
            string text = error.ToText();
            Assert.True(ExcelErrorText.TryParse(text, out ExcelError parsed));
            Assert.Equal(error, parsed);
        }

        Assert.False(ExcelErrorText.TryParse("#BOGUS!", out _));
    }
}
