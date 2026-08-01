using Cellwright.Model;
using Xunit;

namespace Cellwright.Tests.Model;

public sealed class SheetGridTests
{
    [Fact]
    public void SetAndGet_AcrossPageBoundaries()
    {
        var grid = new SheetGrid<CellValue>();
        grid.Set(1, 1, 1.0);
        grid.Set(512, 1, 512.0);   // Last slot of page 0.
        grid.Set(513, 1, 513.0);   // First slot of page 1.
        grid.Set(1_048_576, 16_384, 42.0);

        Assert.Equal(1.0, grid.Get(1, 1).GetNumber());
        Assert.Equal(512.0, grid.Get(512, 1).GetNumber());
        Assert.Equal(513.0, grid.Get(513, 1).GetNumber());
        Assert.Equal(42.0, grid.Get(1_048_576, 16_384).GetNumber());
        Assert.True(grid.Get(100, 100).IsBlank);
        Assert.Equal(1_048_576, grid.MaxRow);
        Assert.Equal(16_384, grid.MaxColumn);
    }

    [Fact]
    public void ClearingWithDefault_RemovesContent()
    {
        var grid = new SheetGrid<CellValue>();
        grid.Set(5, 5, "text");
        grid.Set(5, 5, CellValue.Blank);
        Assert.True(grid.Get(5, 5).IsBlank);
        Assert.False(grid.IsRowOccupied(5));
    }

    [Fact]
    public void ReadRow_FillsBufferAndClearsTail()
    {
        var grid = new SheetGrid<int>();
        grid.Set(1, 2, 20);
        grid.Set(1, 4, 40);

        Span<int> buffer = stackalloc int[6];
        buffer.Fill(-1);
        grid.ReadRow(1, buffer, 6);

        Assert.Equal(0, buffer[0]);
        Assert.Equal(20, buffer[1]);
        Assert.Equal(0, buffer[2]);
        Assert.Equal(40, buffer[3]);
        Assert.Equal(0, buffer[4]);
        Assert.Equal(0, buffer[5]);
    }

    [Fact]
    public void ShiftRows_InsertMovesLaterRowsDownAndLeavesAGap()
    {
        var grid = new SheetGrid<int>();
        for (int row = 1; row <= 4; row++)
            grid.Set(row, 1, row * 10);
        grid.Set(600, 1, 600);   // Second page, so the move crosses a page boundary.

        grid.ShiftRows(2, 2);

        Assert.Equal(10, grid.Get(1, 1));
        Assert.Equal(0, grid.Get(2, 1));
        Assert.Equal(0, grid.Get(3, 1));
        Assert.Equal(20, grid.Get(4, 1));
        Assert.Equal(40, grid.Get(6, 1));
        Assert.Equal(600, grid.Get(602, 1));
        Assert.Equal(602, grid.MaxRow);
    }

    [Fact]
    public void ShiftRows_DeleteClosesTheGap()
    {
        var grid = new SheetGrid<int>();
        for (int row = 1; row <= 5; row++)
            grid.Set(row, 1, row * 10);

        grid.ShiftRows(2, -2);

        Assert.Equal(10, grid.Get(1, 1));
        Assert.Equal(40, grid.Get(2, 1));
        Assert.Equal(50, grid.Get(3, 1));
        Assert.Equal(0, grid.Get(4, 1));
        Assert.Equal(0, grid.Get(5, 1));
    }

    [Fact]
    public void ShiftRows_InsertDropsWhatFallsOffTheGrid()
    {
        var grid = new SheetGrid<int>();
        grid.Set(SheetLimits.MaxRow, 1, 7);

        grid.ShiftRows(1, 1);

        Assert.Equal(0, grid.Get(SheetLimits.MaxRow, 1));
    }

    [Fact]
    public void ShiftColumns_MovesWholeSlices()
    {
        var grid = new SheetGrid<int>();
        for (int column = 1; column <= 4; column++)
            grid.Set(1, column, column);

        grid.ShiftColumns(2, 1);

        Assert.Equal(1, grid.Get(1, 1));
        Assert.Equal(0, grid.Get(1, 2));
        Assert.Equal(2, grid.Get(1, 3));
        Assert.Equal(4, grid.Get(1, 5));

        grid.ShiftColumns(2, -1);

        Assert.Equal(2, grid.Get(1, 2));
        Assert.Equal(4, grid.Get(1, 4));
        Assert.Equal(0, grid.Get(1, 5));
    }

    [Fact]
    public void OutOfBounds_Throws()
    {
        var grid = new SheetGrid<int>();
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.Set(0, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.Set(1, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.Set(SheetLimits.MaxRow + 1, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.Set(1, SheetLimits.MaxColumn + 1, 1));
    }
}
