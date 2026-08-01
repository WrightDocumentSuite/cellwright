using Xunit;

namespace Cellwright.Tests.Primitives;

public sealed class ExcelSerialDateTests
{
    [Theory]
    [InlineData(1900, 1, 1, 1.0)]
    [InlineData(1900, 2, 28, 59.0)]
    [InlineData(1900, 3, 1, 61.0)]
    [InlineData(2023, 3, 15, 45_000.0)]
    [InlineData(9999, 12, 31, 2_958_465.0)]
    public void FromDateTime_1900System(int year, int month, int day, double expected)
    {
        Assert.Equal(expected, ExcelSerialDate.FromDateTime(new DateTime(year, month, day)));
    }

    [Theory]
    [InlineData(1.0, 1900, 1, 1)]
    [InlineData(59.0, 1900, 2, 28)]
    [InlineData(60.0, 1900, 2, 28)] // The fictitious 1900-02-29 maps to the closest real day.
    [InlineData(61.0, 1900, 3, 1)]
    [InlineData(45_000.0, 2023, 3, 15)]
    public void ToDateTime_1900System(double serial, int year, int month, int day)
    {
        Assert.Equal(new DateTime(year, month, day), ExcelSerialDate.ToDateTime(serial));
    }

    [Fact]
    public void FractionalSerial_CarriesTimeOfDay()
    {
        var stamp = new DateTime(2026, 7, 24, 18, 0, 0);
        double serial = ExcelSerialDate.FromDateTime(stamp);
        Assert.Equal(stamp, ExcelSerialDate.ToDateTime(serial));
        Assert.Equal(0.75, serial - Math.Floor(serial), precision: 9);
    }

    [Theory]
    [InlineData(1904, 1, 1, 0.0)]
    [InlineData(2020, 1, 1, 42_369.0)]
    public void Date1904System_RoundTrips(int year, int month, int day, double expected)
    {
        var date = new DateTime(year, month, day);
        double serial = ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Date1904);
        Assert.Equal(expected, serial);
        Assert.Equal(date, ExcelSerialDate.ToDateTime(serial, ExcelDateSystem.Date1904));
    }

    [Fact]
    public void SerialZero_RoundTripsAsExcelsFictitiousDayZero()
    {
        // Time-only cells carry a serial below one; the conversion has to survive the trip
        // in both directions rather than throwing on the way back.
        var timeOnly = new DateTime(1899, 12, 31, 13, 30, 0);
        double serial = ExcelSerialDate.FromDateTime(timeOnly);
        Assert.Equal(0.5625, serial, precision: 9);
        Assert.Equal(timeOnly, ExcelSerialDate.ToDateTime(serial));
    }

    [Fact]
    public void DatesBeforeEpoch_Throw()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.FromDateTime(new DateTime(1899, 12, 30)));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.FromDateTime(new DateTime(1903, 1, 1), ExcelDateSystem.Date1904));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.ToDateTime(-1.0));
    }

    [Fact]
    public void TimeSpan_RoundTrips()
    {
        var duration = new TimeSpan(36, 30, 0);
        double serial = ExcelSerialDate.FromTimeSpan(duration);
        Assert.Equal(1.520833333, serial, precision: 9);
        Assert.Equal(duration, ExcelSerialDate.ToTimeSpan(serial));
    }
}
