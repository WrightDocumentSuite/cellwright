namespace Cellwright.Mapping;

/// <summary>
/// Thrown when sheet data cannot be bound to a row schema: missing headers or
/// a cell value that does not convert to the target property type.
/// </summary>
public sealed class CellMappingException : Exception
{
    /// <summary>Creates the exception with a message.</summary>
    public CellMappingException(string message) : base(message)
    {
    }

    /// <summary>Creates the exception with a message and inner cause.</summary>
    public CellMappingException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
