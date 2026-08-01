namespace Cellwright;

/// <summary>
/// Discriminator of the value stored in a <see cref="CellValue"/>.
/// </summary>
public enum CellValueKind : byte
{
    /// <summary>No value. The default state of a cell.</summary>
    Blank = 0,

    /// <summary>A double-precision number. Excel stores all numerics this way.</summary>
    Number = 1,

    /// <summary>A text value.</summary>
    Text = 2,

    /// <summary>A logical TRUE/FALSE value.</summary>
    Boolean = 3,

    /// <summary>A point in time, stored with full <see cref="System.DateTime"/> tick precision.</summary>
    DateTime = 4,

    /// <summary>A duration, stored with full <see cref="System.TimeSpan"/> tick precision.</summary>
    TimeSpan = 5,

    /// <summary>An <see cref="ExcelError"/> value.</summary>
    Error = 6,
}
