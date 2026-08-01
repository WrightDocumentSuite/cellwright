using Xunit;

namespace Cellwright.Tests.Primitives;

public sealed class CellValueTests
{
    [Fact]
    public void Blank_IsDefault()
    {
        CellValue value = default;
        Assert.Equal(CellValueKind.Blank, value.Kind);
        Assert.True(value.IsBlank);
        Assert.Equal(CellValue.Blank, value);
    }

    [Fact]
    public void Number_RoundTrips()
    {
        CellValue value = 42.5;
        Assert.True(value.IsNumber);
        Assert.Equal(42.5, value.GetNumber());
        Assert.Equal(42.5, (double)value);
    }

    [Fact]
    public void Number_NanAndInfinity_CoercedToNumError()
    {
        Assert.Equal(ExcelError.Number, CellValue.Number(double.NaN).GetError());
        Assert.Equal(ExcelError.Number, CellValue.Number(double.PositiveInfinity).GetError());
        Assert.Equal(ExcelError.Number, CellValue.Number(double.NegativeInfinity).GetError());
    }

    [Fact]
    public void Text_RoundTrips_AndNullBecomesBlank()
    {
        CellValue value = "hello";
        Assert.Equal("hello", value.GetText());

        CellValue fromNull = (string?)null;
        Assert.True(fromNull.IsBlank);
    }

    [Fact]
    public void Text_TooLong_Throws()
    {
        string text = new('x', SheetLimits.MaxCellTextLength + 1);
        Assert.Throws<ArgumentException>(() => CellValue.Text(text));
    }

    [Fact]
    public void DateTime_KeepsTickPrecision()
    {
        var timestamp = new DateTime(2026, 7, 24, 17, 30, 12).AddTicks(1234567);
        CellValue value = timestamp;
        Assert.True(value.IsDateTime);
        Assert.Equal(timestamp, value.GetDateTime());
    }

    [Fact]
    public void TimeSpan_KeepsTickPrecision()
    {
        var duration = new TimeSpan(1, 2, 3, 4, 5) + System.TimeSpan.FromTicks(7);
        CellValue value = duration;
        Assert.Equal(duration, value.GetTimeSpan());
    }

    [Fact]
    public void Boolean_And_Error_RoundTrip()
    {
        CellValue truth = true;
        Assert.True(truth.GetBoolean());

        CellValue error = ExcelError.DivisionByZero;
        Assert.Equal(ExcelError.DivisionByZero, error.GetError());
        Assert.Equal("#DIV/0!", error.ToString());
    }

    [Fact]
    public void WrongKindAccess_Throws()
    {
        CellValue number = 1.0;
        Assert.Throws<InvalidCastException>(() => number.GetText());
        Assert.Throws<InvalidCastException>(() => number.GetBoolean());
        Assert.Throws<InvalidCastException>(() => CellValue.Blank.GetNumber());
    }

    [Fact]
    public void Equality_ComparesKindAndPayload()
    {
        Assert.Equal(CellValue.Number(1), CellValue.Number(1));
        Assert.NotEqual(CellValue.Number(1), CellValue.Boolean(true));
        Assert.Equal(CellValue.Text("a"), CellValue.Text("a"));
        Assert.NotEqual(CellValue.Text("a"), CellValue.Text("A"));
    }

    [Fact]
    public void ToSerialNumber_ConvertsTemporalKinds()
    {
        Assert.Equal(2.5, CellValue.Number(2.5).ToSerialNumber());
        Assert.Equal(1.0, CellValue.Boolean(true).ToSerialNumber());
        Assert.Equal(0.5, CellValue.TimeSpan(System.TimeSpan.FromHours(12)).ToSerialNumber());
        Assert.Equal(45_000.0, CellValue.DateTime(new DateTime(2023, 3, 15)).ToSerialNumber());
        Assert.Throws<InvalidCastException>(() => CellValue.Text("x").ToSerialNumber());
    }

    [Fact]
    public void NullableConversions_ProduceBlank()
    {
        Assert.True(((CellValue)(double?)null).IsBlank);
        Assert.True(((CellValue)(DateTime?)null).IsBlank);
        Assert.Equal(3.0, ((CellValue)(int?)3).GetNumber());
    }
}
