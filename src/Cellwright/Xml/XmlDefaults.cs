using System.Xml;

namespace Cellwright.Xml;

/// <summary>
/// Hardened <see cref="XmlReader"/> settings shared by all parsing code paths.
/// </summary>
internal static class XmlDefaults
{
    /// <summary>Settings for synchronous parsing of buffered part content.</summary>
    public static readonly XmlReaderSettings ReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        IgnoreWhitespace = true,
        CloseInput = false,
    };

    /// <summary>Settings for asynchronous streaming parsing of large parts.</summary>
    public static readonly XmlReaderSettings AsyncReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        IgnoreWhitespace = true,
        CloseInput = false,
        Async = true,
    };
}
