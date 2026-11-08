using System.Security.Cryptography;
using System.Xml;
using Cellwright;
using Cellwright.Xls;
using SharpFuzz;

// Do not execute instrumented Cellwright code before LibFuzzer.Run initializes its coverage
// channel. Even an options constructor contains probes after SharpFuzz instrumentation.
XlsLoadOptions? options = null;

Fuzzer.LibFuzzer.Run(data =>
{
    options ??= CreateOptions();
    using var stream = new MemoryStream(data.ToArray(), writable: false);
    try
    {
        Spreadsheet.LoadAsync(stream, options).AsTask().GetAwaiter().GetResult();
    }
    catch (Exception exception) when (IsRejectedInput(exception))
    {
        // Invalid, unsupported and over-budget documents are normal fuzz outcomes. Anything
        // else reaches libFuzzer as a crash and is retained as a reproducer.
    }
});

static XlsLoadOptions CreateOptions() => new()
{
    Password = "password",
    LeaveOpen = true,
    Budget = new DocumentLoadBudget
    {
        MaxInputBytes = 1024 * 1024,
        MaxPackageParts = 2_048,
        MaxInflatedBytes = 32 * 1024 * 1024,
        MaxPartBytes = 8 * 1024 * 1024,
        MaxXmlCharactersPerPart = 8 * 1024 * 1024,
        MaxXmlNodes = 1_000_000,
        MaxXmlDepth = 128,
        MaxCells = 250_000,
        MaxSharedStrings = 100_000,
        MaxCellTextCharacters = 64 * 1024,
        MaxTotalCellTextCharacters = 8 * 1024 * 1024,
        MaxMediaBytes = 4 * 1024 * 1024,
        MaxTotalMediaBytes = 8 * 1024 * 1024,
        MaxEmbeddedObjectBytes = 4 * 1024 * 1024,
        MaxEmbeddedObjects = 64,
    },
};

static bool IsRejectedInput(Exception exception) => exception switch
{
    XlsFormatException or XlsxFormatException or DocumentLoadLimitException => true,
    InvalidDataException or IOException or NotSupportedException => true,
    XmlException or CryptographicException => true,
    ArgumentException when exception.GetType() == typeof(ArgumentException) => true,
    _ => false,
};
