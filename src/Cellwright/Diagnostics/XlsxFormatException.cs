namespace Cellwright;

/// <summary>
/// Thrown when an xlsx package is structurally broken beyond what recovery-oriented
/// loading can repair (missing required parts, malformed XML, invalid zip data).
/// </summary>
public sealed class XlsxFormatException : Exception
{
    /// <summary>Creates the exception with a message.</summary>
    public XlsxFormatException(string message) : base(message)
    {
    }

    /// <summary>Creates the exception with a message and inner cause.</summary>
    public XlsxFormatException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
