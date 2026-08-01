using Xunit;

namespace Cellwright.Tests.Primitives;

public sealed class RangeRefTests
{
    [Fact]
    public void Parse_TwoCorners()
    {
        var range = RangeRef.Parse("A1:C5");
        Assert.Equal(new CellRef(1, 1), range.Start);
        Assert.Equal(new CellRef(5, 3), range.End);
        Assert.Equal(5, range.RowCount);
        Assert.Equal(3, range.ColumnCount);
        Assert.Equal(15, range.CellCount);
    }

    [Fact]
    public void Parse_SingleCell()
    {
        var range = RangeRef.Parse("B2");
        Assert.True(range.IsSingleCell);
        Assert.Equal("B2", range.ToString());
    }

    [Fact]
    public void Constructor_NormalizesCorners()
    {
        var range = new RangeRef(CellRef.Parse("C5"), CellRef.Parse("A1"));
        Assert.Equal("A1:C5", range.ToString());
    }

    [Theory]
    [InlineData("A1:B2", "B2", true)]
    [InlineData("A1:B2", "C1", false)]
    public void Contains_Works(string rangeText, string cellText, bool expected)
    {
        Assert.Equal(expected, RangeRef.Parse(rangeText).Contains(CellRef.Parse(cellText)));
    }

    [Theory]
    [InlineData("A1:C3", "B2:D4", true)]
    [InlineData("A1:B2", "C3:D4", false)]
    [InlineData("A1:A10", "A5", true)]
    public void Intersects_Works(string first, string second, bool expected)
    {
        Assert.Equal(expected, RangeRef.Parse(first).Intersects(RangeRef.Parse(second)));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1:")]
    [InlineData(":B2")]
    [InlineData("A1:B2:C3")]
    public void Parse_Invalid_Throws(string text)
    {
        Assert.Throws<FormatException>(() => RangeRef.Parse(text));
    }

    [Fact]
    public void TryFormat_Utf8_MatchesText()
    {
        var range = RangeRef.Parse("A1:XFD1048576");
        Span<byte> buffer = stackalloc byte[RangeRef.MaxTextLength];
        Assert.True(range.TryFormat(buffer, out int written, default, null));
        Assert.Equal("A1:XFD1048576"u8, buffer[..written]);
    }
}
