namespace Cellwright.Xls.Reading;

/// <summary>
/// Reinterprets a number as the date or duration its format says it is. A spreadsheet has no
/// date type: a date is a number with a date format on it, so the format table is the only
/// thing that says one apart from the other.
/// </summary>
internal static class TemporalValue
{
    /// <summary>
    /// Wraps a serial number as a date, or leaves it a number when it does not name one. A file
    /// may perfectly well carry a negative number under a date format, and turning that into an
    /// exception would fail the load over a cell Excel shows as <c>######</c>.
    /// </summary>
    public static CellValue AsDate(double serial, ExcelDateSystem dateSystem)
    {
        if (double.IsNaN(serial) || serial < 0 || serial > MaxSerial)
            return CellValue.Number(serial);

        try
        {
            return CellValue.DateTime(ExcelSerialDate.ToDateTime(serial, dateSystem));
        }
        catch (ArgumentOutOfRangeException)
        {
            return CellValue.Number(serial);
        }
    }

    /// <summary>Wraps a serial number as an elapsed duration.</summary>
    public static CellValue AsDuration(double serial)
    {
        if (double.IsNaN(serial) || Math.Abs(serial) > MaxDurationDays)
            return CellValue.Number(serial);

        return CellValue.TimeSpan(ExcelSerialDate.ToTimeSpan(serial));
    }

    /// <summary>Serial of 9999-12-31, the last date Excel represents.</summary>
    private const double MaxSerial = 2_958_465.0;

    /// <summary>Days a <see cref="TimeSpan"/> can hold, near enough.</summary>
    private const double MaxDurationDays = 10_675_199.0;
}
