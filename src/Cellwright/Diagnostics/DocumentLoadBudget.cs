using System.Buffers;
using System.Globalization;

namespace Cellwright;

/// <summary>
/// Bounds the memory and parser work one spreadsheet import may spend on caller-controlled input.
/// </summary>
/// <remarks>
/// Limits are inclusive and shared by every reader participating in one load. The defaults are
/// deliberately generous for real workbooks, but finite; a caller may raise individual ceilings
/// for unusually large trusted files.
/// </remarks>
public sealed record DocumentLoadBudget
{
    private const long MiB = 1024L * 1024L;

    /// <summary>The shared default budget.</summary>
    public static DocumentLoadBudget Default { get; } = new();

    /// <summary>No ceilings, for compatibility when a caller explicitly trusts the source.</summary>
    public static DocumentLoadBudget Unlimited { get; } = new()
    {
        MaxInputBytes = long.MaxValue,
        MaxPackageParts = int.MaxValue,
        MaxInflatedBytes = long.MaxValue,
        MaxPartBytes = long.MaxValue,
        MaxXmlCharactersPerPart = long.MaxValue,
        MaxXmlNodes = long.MaxValue,
        MaxXmlDepth = int.MaxValue,
        MaxCells = long.MaxValue,
        MaxSharedStrings = int.MaxValue,
        MaxCellTextCharacters = int.MaxValue,
        MaxTotalCellTextCharacters = long.MaxValue,
        MaxMediaBytes = long.MaxValue,
        MaxTotalMediaBytes = long.MaxValue,
        MaxEmbeddedObjectBytes = long.MaxValue,
        MaxEmbeddedObjects = int.MaxValue,
    };

    /// <summary>Largest source file or stream. The default is 512 MiB.</summary>
    public long MaxInputBytes { get; init; } = 512 * MiB;

    /// <summary>Largest number of OPC parts or CFB directory entries. The default is 16,384.</summary>
    public int MaxPackageParts { get; init; } = 16_384;

    /// <summary>Largest total uncompressed package/container payload. The default is 2 GiB.</summary>
    public long MaxInflatedBytes { get; init; } = 2L * 1024 * MiB;

    /// <summary>Largest uncompressed package part or container stream. The default is 512 MiB.</summary>
    public long MaxPartBytes { get; init; } = 512 * MiB;

    /// <summary>Largest XML part in decoded characters. The default is 512 Mi characters.</summary>
    public long MaxXmlCharactersPerPart { get; init; } = 512 * MiB;

    /// <summary>Largest total number of XML nodes read from one package. The default is 64 million.</summary>
    public long MaxXmlNodes { get; init; } = 64_000_000;

    /// <summary>Largest XML element nesting depth. The default is 256.</summary>
    public int MaxXmlDepth { get; init; } = 256;

    /// <summary>Largest number of cells materialized or streamed. The default is ten million.</summary>
    public long MaxCells { get; init; } = 10_000_000;

    /// <summary>Largest number of shared-string items. The default is four million.</summary>
    public int MaxSharedStrings { get; init; } = 4_000_000;

    /// <summary>Largest stored cell/shared-string source item. The default is one Mi characters.</summary>
    public int MaxCellTextCharacters { get; init; } = 1 * 1024 * 1024;

    /// <summary>Largest total source text stored in cells and shared strings. The default is 256 Mi characters.</summary>
    public long MaxTotalCellTextCharacters { get; init; } = 256 * MiB;

    /// <summary>Largest decoded image, audio or video part. The default is 64 MiB.</summary>
    public long MaxMediaBytes { get; init; } = 64 * MiB;

    /// <summary>Largest total decoded media payload. The default is 512 MiB.</summary>
    public long MaxTotalMediaBytes { get; init; } = 512 * MiB;

    /// <summary>Largest embedded object payload. The default is 128 MiB.</summary>
    public long MaxEmbeddedObjectBytes { get; init; } = 128 * MiB;

    /// <summary>Largest number of embedded objects. The default is 4,096.</summary>
    public int MaxEmbeddedObjects { get; init; } = 4_096;

    /// <summary>Checks that every configured ceiling is positive.</summary>
    public void Validate()
    {
        Positive(MaxInputBytes, nameof(MaxInputBytes));
        Positive(MaxPackageParts, nameof(MaxPackageParts));
        Positive(MaxInflatedBytes, nameof(MaxInflatedBytes));
        Positive(MaxPartBytes, nameof(MaxPartBytes));
        Positive(MaxXmlCharactersPerPart, nameof(MaxXmlCharactersPerPart));
        Positive(MaxXmlNodes, nameof(MaxXmlNodes));
        Positive(MaxXmlDepth, nameof(MaxXmlDepth));
        Positive(MaxCells, nameof(MaxCells));
        Positive(MaxSharedStrings, nameof(MaxSharedStrings));
        Positive(MaxCellTextCharacters, nameof(MaxCellTextCharacters));
        Positive(MaxTotalCellTextCharacters, nameof(MaxTotalCellTextCharacters));
        Positive(MaxMediaBytes, nameof(MaxMediaBytes));
        Positive(MaxTotalMediaBytes, nameof(MaxTotalMediaBytes));
        Positive(MaxEmbeddedObjectBytes, nameof(MaxEmbeddedObjectBytes));
        Positive(MaxEmbeddedObjects, nameof(MaxEmbeddedObjects));
    }

    private static void Positive(long value, string name)
    {
        if (value < 1)
            throw new ArgumentOutOfRangeException(name, value, "A document load limit must be positive.");
    }
}

/// <summary>Thrown when an import exceeds one configured <see cref="DocumentLoadBudget"/> ceiling.</summary>
public sealed class DocumentLoadLimitException : Exception
{
    /// <summary>Creates an exception with stable machine-readable limit information.</summary>
    public DocumentLoadLimitException(string limitName, long limit, long observed)
        : base(FormattableString.Invariant(
            $"Document load limit '{limitName}' exceeded (limit: {limit.ToString(CultureInfo.InvariantCulture)}, observed: {observed.ToString(CultureInfo.InvariantCulture)})."))
    {
        ArgumentException.ThrowIfNullOrEmpty(limitName);
        LimitName = limitName;
        Limit = limit;
        Observed = observed;
    }

    /// <summary>Name of the <see cref="DocumentLoadBudget"/> property that was exceeded.</summary>
    public string LimitName { get; }

    /// <summary>Configured inclusive ceiling.</summary>
    public long Limit { get; }

    /// <summary>Value observed when the reader stopped.</summary>
    public long Observed { get; }
}

/// <summary>Mutable counters shared by all readers involved in one spreadsheet load.</summary>
internal sealed class DocumentLoadBudgetState
{
    private long _packageParts;
    private long _inflatedBytes;
    private long _xmlNodes;
    private long _cells;
    private long _sharedStrings;
    private long _cellTextCharacters;
    private long _mediaBytes;
    private long _embeddedObjects;

    public DocumentLoadBudgetState(DocumentLoadBudget budget)
    {
        Budget = budget ?? throw new ArgumentNullException(nameof(budget));
        budget.Validate();
    }

    public DocumentLoadBudget Budget { get; }

    public static void Ensure(string limitName, long limit, long observed)
    {
        if (observed > limit)
            throw new DocumentLoadLimitException(limitName, limit, observed);
    }

    public void AddPackageParts(long count)
    {
        _packageParts = AddWithoutOverflow(_packageParts, count);
        Ensure(nameof(DocumentLoadBudget.MaxPackageParts), Budget.MaxPackageParts, _packageParts);
    }

    public void AddInflatedBytes(long bytes)
    {
        _inflatedBytes = AddWithoutOverflow(_inflatedBytes, bytes);
        Ensure(nameof(DocumentLoadBudget.MaxInflatedBytes), Budget.MaxInflatedBytes, _inflatedBytes);
    }

    public void AddXmlNode(int depth)
    {
        Ensure(nameof(DocumentLoadBudget.MaxXmlDepth), Budget.MaxXmlDepth, depth);
        _xmlNodes = AddWithoutOverflow(_xmlNodes, 1);
        Ensure(nameof(DocumentLoadBudget.MaxXmlNodes), Budget.MaxXmlNodes, _xmlNodes);
    }

    public void AddCell()
    {
        _cells = AddWithoutOverflow(_cells, 1);
        Ensure(nameof(DocumentLoadBudget.MaxCells), Budget.MaxCells, _cells);
    }

    public void AddSharedString()
    {
        AddSharedStrings(1);
    }

    public void AddSharedStrings(int count)
    {
        _sharedStrings = AddWithoutOverflow(_sharedStrings, count);
        Ensure(nameof(DocumentLoadBudget.MaxSharedStrings), Budget.MaxSharedStrings, _sharedStrings);
    }

    public void AddCellText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        long itemCharacters = 0;
        AddCellText(text.Length, ref itemCharacters);
    }

    public void EnsureCellTextLength(long characters) =>
        Ensure(nameof(DocumentLoadBudget.MaxCellTextCharacters), Budget.MaxCellTextCharacters, characters);

    public void AddCellText(int characters, ref long itemCharacters)
    {
        itemCharacters = AddWithoutOverflow(itemCharacters, characters);
        Ensure(nameof(DocumentLoadBudget.MaxCellTextCharacters), Budget.MaxCellTextCharacters, itemCharacters);
        _cellTextCharacters = AddWithoutOverflow(_cellTextCharacters, characters);
        Ensure(
            nameof(DocumentLoadBudget.MaxTotalCellTextCharacters),
            Budget.MaxTotalCellTextCharacters,
            _cellTextCharacters);
    }

    public void EnsureMedia(long bytes)
    {
        Ensure(nameof(DocumentLoadBudget.MaxMediaBytes), Budget.MaxMediaBytes, bytes);
        long observed = AddWithoutOverflow(_mediaBytes, bytes);
        Ensure(nameof(DocumentLoadBudget.MaxTotalMediaBytes), Budget.MaxTotalMediaBytes, observed);
    }

    public void AddMedia(long bytes)
    {
        EnsureMedia(bytes);
        long observed = AddWithoutOverflow(_mediaBytes, bytes);
        _mediaBytes = observed;
    }

    public void AddMediaDelta(long itemBytes, long additionalBytes)
    {
        Ensure(nameof(DocumentLoadBudget.MaxMediaBytes), Budget.MaxMediaBytes, itemBytes);
        long observed = AddWithoutOverflow(_mediaBytes, additionalBytes);
        Ensure(nameof(DocumentLoadBudget.MaxTotalMediaBytes), Budget.MaxTotalMediaBytes, observed);
        _mediaBytes = observed;
    }

    public long MaximumNextMediaBytes =>
        Math.Min(Budget.MaxMediaBytes, Math.Max(0, Budget.MaxTotalMediaBytes - _mediaBytes));

    public void AddEmbeddedObject(long bytes)
    {
        Ensure(nameof(DocumentLoadBudget.MaxEmbeddedObjectBytes), Budget.MaxEmbeddedObjectBytes, bytes);
        _embeddedObjects = AddWithoutOverflow(_embeddedObjects, 1);
        Ensure(nameof(DocumentLoadBudget.MaxEmbeddedObjects), Budget.MaxEmbeddedObjects, _embeddedObjects);
    }

    public static long AddWithoutOverflow(long left, long right) =>
        right > long.MaxValue - left ? long.MaxValue : left + right;
}

/// <summary>Bounded whole-input reads used by binary formats and non-seekable package streams.</summary>
internal static class DocumentInput
{
    public static async ValueTask<byte[]> ReadBytesAsync(
        Stream stream, DocumentLoadBudget budget, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(budget);
        budget.Validate();

        if (stream.CanSeek)
        {
            long remaining = Math.Max(0, stream.Length - stream.Position);
            DocumentLoadBudgetState.Ensure(
                nameof(DocumentLoadBudget.MaxInputBytes), budget.MaxInputBytes, remaining);
        }

        using var content = new MemoryStream((int)Math.Min(budget.MaxInputBytes, 81920));
        byte[] buffer = ArrayPool<byte>.Shared.Rent(81920);
        try
        {
            long total = 0;
            while (true)
            {
                int read = await stream.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
                if (read == 0)
                    return content.ToArray();

                long observed = DocumentLoadBudgetState.AddWithoutOverflow(total, read);
                DocumentLoadBudgetState.Ensure(
                    nameof(DocumentLoadBudget.MaxInputBytes), budget.MaxInputBytes, observed);
                await content.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                total = observed;
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
}
