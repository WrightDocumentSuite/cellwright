namespace Cellwright.Xls;

/// <summary>
/// A legacy workbook could not be read. Thrown only when the file is broken past the point of
/// guessing; anything recoverable arrives as a <see cref="WorkbookWarning"/> instead.
/// </summary>
public sealed class XlsFormatException : Exception
{
    /// <summary>Creates the exception with a message.</summary>
    public XlsFormatException(string message) : base(message)
    {
    }

    /// <summary>Creates the exception with a message and the failure underneath it.</summary>
    public XlsFormatException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
