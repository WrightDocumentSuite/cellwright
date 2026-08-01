namespace Cellwright;

/// <summary>
/// The date system a workbook uses to map serial numbers to calendar dates.
/// </summary>
public enum ExcelDateSystem
{
    /// <summary>Serial 1 is 1900-01-01. Default for xlsx. Inherits the fictitious 1900-02-29 from Lotus 1-2-3.</summary>
    Date1900 = 0,

    /// <summary>Serial 0 is 1904-01-01. Legacy Mac option, no leap-year bug.</summary>
    Date1904 = 1,
}

/// <summary>
/// Precise conversions between <see cref="DateTime"/>/<see cref="TimeSpan"/> and Excel serial date numbers.
/// Handles the fictitious 1900-02-29 (serial 60 in the 1900 system) explicitly instead of
/// silently shifting January/February 1900 dates the way naive OADate round-trips do.
/// </summary>
public static class ExcelSerialDate
{
    private const double Serial1900Cutover = 61.0; // 1900-03-01 in the 1900 system: serial == OADate from here on.
    private const double Date1904Offset = 1462.0;  // Days from the OADate epoch (1899-12-30) to 1904-01-01.

    // Serial 0 of the 1900 system is Excel's fictitious "1900-01-00", which lands on
    // 1899-12-31 in real time. Time-only cells use it, so the conversion must accept it or a
    // value like 13:30 formatted as a date-time would fail to convert back.
    private static readonly DateTime MinDate1900 = new(1899, 12, 31);
    private static readonly DateTime MinDate1904 = new(1904, 1, 1);

    /// <summary>
    /// Converts a date-time to its Excel serial number.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The date precedes the epoch of the chosen date system.</exception>
    public static double FromDateTime(DateTime value, ExcelDateSystem system = ExcelDateSystem.Date1900)
    {
        if (system == ExcelDateSystem.Date1904)
        {
            if (value < MinDate1904)
                throw new ArgumentOutOfRangeException(nameof(value), value, "The 1904 date system cannot represent dates before 1904-01-01.");
            return value.ToOADate() - Date1904Offset;
        }

        if (value < MinDate1900)
            throw new ArgumentOutOfRangeException(nameof(value), value, "The 1900 date system cannot represent dates before 1899-12-31.");

        double oa = value.ToOADate();
        // Before 1900-03-01 Excel serials run one behind OADate because Excel believes 1900-02-29 existed.
        return oa < Serial1900Cutover ? oa - 1.0 : oa;
    }

    /// <summary>
    /// Converts an Excel serial number to a <see cref="DateTime"/>.
    /// Serial 60 of the 1900 system (the fictitious 1900-02-29) maps to 1900-02-28.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The serial number is negative or too large.</exception>
    public static DateTime ToDateTime(double serial, ExcelDateSystem system = ExcelDateSystem.Date1900)
    {
        if (double.IsNaN(serial) || serial < 0.0)
            throw new ArgumentOutOfRangeException(nameof(serial), serial, "Excel date serial numbers must be non-negative.");

        double oa = system == ExcelDateSystem.Date1904
            ? serial + Date1904Offset
            : serial < 60.0 ? serial + 1.0 : serial;

        return DateTime.FromOADate(oa);
    }

    /// <summary>Converts a duration to the serial representation Excel uses for elapsed time (total days).</summary>
    public static double FromTimeSpan(TimeSpan value) => value.TotalDays;

    /// <summary>Converts a serial number interpreted as elapsed time back to a duration.</summary>
    public static TimeSpan ToTimeSpan(double serial)
    {
        // TimeSpan.FromDays truncates towards zero, which costs a tick on values that only
        // just miss a tick boundary — enough to make a duration fail to survive a round trip.
        double ticks = Math.Round(serial * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
        if (double.IsNaN(ticks) || ticks < long.MinValue || ticks > long.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(serial), serial, "The duration does not fit a TimeSpan.");

        return new TimeSpan((long)ticks);
    }
}
