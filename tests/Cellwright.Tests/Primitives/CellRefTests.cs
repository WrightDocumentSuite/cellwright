using Xunit;

namespace Cellwright.Tests.Primitives;

public sealed class CellRefTests
{
    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("B7", 7, 2)]
    [InlineData("Z99", 99, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("AB12", 12, 28)]
    [InlineData("XFD1048576", 1_048_576, 16_384)]
    [InlineData("$C$12", 12, 3)]
    [InlineData("$C12", 12, 3)]
    [InlineData("C$12", 12, 3)]
    [InlineData("ab12", 12, 28)]
    public void Parse_ValidReferences(string text, int row, int column)
    {
        var cell = CellRef.Parse(text);
        Assert.Equal(row, cell.Row);
        Assert.Equal(column, cell.Column);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("A-1")]
    [InlineData("A1B")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("AAAA1")]
    [InlineData("A 1")]
    public void Parse_InvalidReferences_Throws(string text)
    {
        Assert.False(CellRef.TryParse(text, null, out _));
        Assert.Throws<FormatException>(() => CellRef.Parse(text));
    }

    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(12, 28, "AB12")]
    [InlineData(1_048_576, 16_384, "XFD1048576")]
    public void ToString_FormatsA1(int row, int column, string expected)
    {
        Assert.Equal(expected, new CellRef(row, column).ToString());
    }

    [Fact]
    public void TryFormat_Utf8_MatchesText()
    {
        var cell = new CellRef(1_048_576, 16_384);
        Span<byte> buffer = stackalloc byte[CellRef.MaxTextLength];
        Assert.True(cell.TryFormat(buffer, out int written, default, null));
        Assert.Equal("XFD1048576"u8, buffer[..written]);
    }

    [Fact]
    public void IsValid_ChecksGridBounds()
    {
        Assert.True(new CellRef(1, 1).IsValid);
        Assert.True(new CellRef(SheetLimits.MaxRow, SheetLimits.MaxColumn).IsValid);
        Assert.False(new CellRef(0, 1).IsValid);
        Assert.False(new CellRef(1, SheetLimits.MaxColumn + 1).IsValid);
    }
}
